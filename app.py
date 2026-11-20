"""Sistema de control de ventas - Supermercado de Repostería (Sucre, Bolivia).

Ejecutar:  .venv/bin/python app.py   y abrir http://127.0.0.1:5000
"""
import io
import math
import os
import random
import re
import sqlite3
import secrets
import tempfile
from datetime import datetime, date, timedelta
from functools import wraps
from zoneinfo import ZoneInfo

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook
from openpyxl.styles import Font
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("REPOSTERIA_DB", os.path.join(BASE_DIR, "reposteria.db"))
# Con DATABASE_URL (Vercel + Neon) se usa Postgres; sin ella, el archivo SQLite local.
DATABASE_URL = os.environ.get("DATABASE_URL") or os.environ.get("POSTGRES_URL")
ES_PG = bool(DATABASE_URL)
EN_VERCEL = bool(os.environ.get("VERCEL"))
TZ = ZoneInfo("America/La_Paz")

if ES_PG:
    import psycopg
    ErrorBD, ErrorIntegridad = psycopg.Error, psycopg.IntegrityError
else:
    ErrorBD, ErrorIntegridad = sqlite3.Error, sqlite3.IntegrityError

METODOS_PAGO = {
    "efectivo": "Efectivo",
    "qr": "QR Simple",
    "tarjeta": "Tarjeta",
    "transferencia": "Transferencia",
}
ESTADOS_ENVIO = {
    "pendiente": "Pendiente",
    "en_camino": "En camino",
    "entregado": "Entregado",
    "cancelado": "Cancelado",
}
CATEGORIAS = ["Harinas y premezclas", "Azúcares y endulzantes", "Chocolates y coberturas",
              "Decoración", "Moldes y utensilios", "Lácteos y huevos", "Esencias y colorantes",
              "Tortas y postres", "Empaques y descartables", "Otros"]

app = Flask(__name__)
def clave_local():
    ruta = os.path.join(BASE_DIR, ".secret_key")
    try:
        with open(ruta, "x") as archivo:
            archivo.write(secrets.token_hex(32))
        os.chmod(ruta, 0o600)
    except FileExistsError:
        pass
    with open(ruta) as archivo:
        return archivo.read().strip()


if EN_VERCEL and not os.environ.get("SECRET_KEY"):
    raise RuntimeError("Falta la variable de entorno SECRET_KEY en Vercel.")
app.config.update(SECRET_KEY=os.environ.get("SECRET_KEY") or clave_local(),
                  SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  SESSION_COOKIE_SECURE=EN_VERCEL,
                  MAX_CONTENT_LENGTH=2 * 1024 * 1024)


def csrf_token():
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(32)
    return session["csrf_token"]


app.jinja_env.globals["csrf_token"] = csrf_token


@app.before_request
def proteger_escrituras():
    if request.method == "POST":
        token = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token", "")
        if not secrets.compare_digest(str(token), session.get("csrf_token", "") or "sin-token"):
            abort(400, description="Formulario vencido. Recargá la página e intentá nuevamente.")
        # Serializa escrituras antes de leer stock o estados desde varias computadoras.
        bloquear_escrituras(get_db())



# ---------------------------------------------------------------- utilidades
def ahora():
    return datetime.now(TZ)


def hoy():
    return ahora().date()


@app.template_filter("bs")
def formato_bs(valor):
    valor = float(valor or 0)
    entero, dec = f"{valor:,.2f}".split(".")
    return f"Bs {entero.replace(',', '.')},{dec}"


@app.template_filter("fecha")
def formato_fecha(valor):
    if not valor:
        return "—"
    if isinstance(valor, str):
        try:
            valor = date.fromisoformat(valor) if len(valor) == 10 else datetime.fromisoformat(valor)
        except ValueError:
            return valor
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y %H:%M")
    return valor.strftime("%d/%m/%Y")


@app.context_processor
def contexto_global():
    return {"METODOS_PAGO": METODOS_PAGO, "ESTADOS_ENVIO": ESTADOS_ENVIO,
            "usuario": g.get("usuario"), "config": obtener_config()}


# ---------------------------------------------------------------- base de datos
class Fila(dict):
    """Fila accesible por nombre (fila["total"]) y por posición (fila[0]), igual en SQLite y Postgres."""

    def __init__(self, columnas, valores):
        super().__init__(zip(columnas, valores))
        self._valores = tuple(valores)

    def __getitem__(self, clave):
        if isinstance(clave, int):
            return self._valores[clave]
        return dict.__getitem__(self, clave)


def _fila_sqlite(cursor, valores):
    return Fila([d[0] for d in cursor.description], valores)


def _fila_pg(cursor):
    columnas = [d.name for d in cursor.description] if cursor.description else []
    return lambda valores: Fila(columnas, valores)


class ConexionPG:
    """Envoltura mínima para usar las mismas consultas con «?» que en SQLite."""

    def __init__(self):
        # prepare_threshold=None: compatible con el pooler de Neon (PgBouncer).
        self.conexion = psycopg.connect(DATABASE_URL, row_factory=_fila_pg,
                                        prepare_threshold=None, connect_timeout=15)

    def execute(self, sql, params=None):
        return self.conexion.execute(sql.replace("?", "%s"), tuple(params) if params else None)

    def commit(self):
        self.conexion.commit()

    def rollback(self):
        self.conexion.rollback()

    def close(self):
        self.conexion.close()


def conectar():
    if ES_PG:
        return ConexionPG()
    db = sqlite3.connect(DB_PATH, timeout=20)
    db.row_factory = _fila_sqlite
    db.execute("PRAGMA foreign_keys = ON")
    return db


def bloquear_escrituras(db):
    if ES_PG:
        db.execute("SELECT pg_advisory_xact_lock(4242)")
    else:
        db.execute("BEGIN IMMEDIATE")


def get_db():
    if "db" not in g:
        g.db = conectar()
    return g.db


@app.teardown_appcontext
def cerrar_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def nid():
    """Negocio del usuario conectado. Todas las consultas se filtran por este valor."""
    return g.usuario["negocio_id"]


ESQUEMA = """
CREATE TABLE IF NOT EXISTS negocios (
    id INTEGER PRIMARY KEY,
    nombre TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY,
    negocio_id INTEGER NOT NULL REFERENCES negocios(id),
    usuario TEXT UNIQUE NOT NULL,
    nombre TEXT NOT NULL,
    rol TEXT NOT NULL CHECK (rol IN ('ventas','envios')),
    clave_hash TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS config (
    negocio_id INTEGER NOT NULL REFERENCES negocios(id),
    clave TEXT NOT NULL,
    valor TEXT NOT NULL,
    PRIMARY KEY (negocio_id, clave)
);
CREATE TABLE IF NOT EXISTS productos (
    id INTEGER PRIMARY KEY,
    negocio_id INTEGER NOT NULL REFERENCES negocios(id),
    codigo TEXT,
    nombre TEXT NOT NULL,
    categoria TEXT NOT NULL,
    precio REAL NOT NULL CHECK (precio >= 0),
    stock INTEGER NOT NULL DEFAULT 0,
    stock_minimo INTEGER NOT NULL DEFAULT 5,
    fecha_vencimiento TEXT,
    activo INTEGER NOT NULL DEFAULT 1,
    UNIQUE (negocio_id, codigo)
);
CREATE TABLE IF NOT EXISTS ventas (
    id INTEGER PRIMARY KEY,
    negocio_id INTEGER NOT NULL REFERENCES negocios(id),
    fecha TEXT NOT NULL,
    usuario_id INTEGER NOT NULL REFERENCES usuarios(id),
    cliente_nombre TEXT NOT NULL DEFAULT 'S/N',
    cliente_nit TEXT NOT NULL DEFAULT '0',
    subtotal REAL NOT NULL,
    costo_envio REAL NOT NULL DEFAULT 0,
    total REAL NOT NULL,
    metodo_pago TEXT NOT NULL,
    monto_recibido REAL,
    cambio REAL,
    referencia_pago TEXT,
    anulada INTEGER NOT NULL DEFAULT 0,
    motivo_anulacion TEXT,
    demo INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS venta_items (
    id INTEGER PRIMARY KEY,
    venta_id INTEGER NOT NULL REFERENCES ventas(id),
    producto_id INTEGER NOT NULL REFERENCES productos(id),
    nombre TEXT NOT NULL,
    cantidad INTEGER NOT NULL CHECK (cantidad > 0),
    precio_unitario REAL NOT NULL,
    total REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS envios (
    id INTEGER PRIMARY KEY,
    venta_id INTEGER UNIQUE NOT NULL REFERENCES ventas(id),
    destinatario TEXT NOT NULL,
    telefono TEXT NOT NULL,
    direccion TEXT NOT NULL,
    referencia TEXT,
    fecha_programada TEXT,
    estado TEXT NOT NULL DEFAULT 'pendiente',
    observaciones TEXT,
    repartidor_id INTEGER REFERENCES usuarios(id),
    actualizado TEXT
);
CREATE TABLE IF NOT EXISTS movimientos_stock (
    id INTEGER PRIMARY KEY,
    fecha TEXT NOT NULL,
    producto_id INTEGER NOT NULL REFERENCES productos(id),
    cantidad INTEGER NOT NULL,
    motivo TEXT NOT NULL,
    usuario_id INTEGER REFERENCES usuarios(id)
);
CREATE INDEX IF NOT EXISTS idx_productos_negocio ON productos (negocio_id);
CREATE INDEX IF NOT EXISTS idx_ventas_negocio_fecha ON ventas (negocio_id, fecha);
CREATE INDEX IF NOT EXISTS idx_items_venta ON venta_items (venta_id);
"""


def esquema_pg():
    sql = ESQUEMA.replace("id INTEGER PRIMARY KEY", "id SERIAL PRIMARY KEY")
    return re.sub(r"\bREAL\b", "DOUBLE PRECISION", sql)


CONFIG_INICIAL = {
    "negocio_nombre": "Supermercado Repostería Sucre",
    "negocio_nit": "0000000000",
    "negocio_direccion": "Sucre, Chuquisaca - Bolivia",
    "negocio_telefono": "",
    "tarifa_envio": "15",
    "dias_aviso_vencimiento": "7",
}
# Nombre que se muestra antes de iniciar sesión (en el sitio compartido no es de ningún negocio).
NOMBRE_SITIO = os.environ.get("NOMBRE_SITIO") or (
    "Sistema de Repostería" if ES_PG else CONFIG_INICIAL["negocio_nombre"])

PRODUCTOS_EJEMPLO = [
    ("HAR-001", "Harina 0000 x 1 kg", "Harinas y premezclas", 12.0, 40, 10, 180),
    ("HAR-002", "Premezcla bizcochuelo vainilla 500 g", "Harinas y premezclas", 18.5, 25, 5, 120),
    ("AZU-001", "Azúcar impalpable x 1 kg", "Azúcares y endulzantes", 14.0, 30, 8, 365),
    ("CHO-001", "Cobertura chocolate semiamargo 1 kg", "Chocolates y coberturas", 65.0, 12, 4, 200),
    ("CHO-002", "Chispas de chocolate 250 g", "Chocolates y coberturas", 16.0, 3, 5, 5),
    ("DEC-001", "Grageas de colores 100 g", "Decoración", 7.5, 50, 10, 300),
    ("DEC-002", "Fondant blanco 1 kg", "Decoración", 45.0, 8, 3, 90),
    ("MOL-001", "Molde redondo desmontable 24 cm", "Moldes y utensilios", 55.0, 10, 2, None),
    ("LAC-001", "Crema de leche 1 L", "Lácteos y huevos", 28.0, 15, 6, 4),
    ("ESE-001", "Esencia de vainilla 100 ml", "Esencias y colorantes", 10.0, 20, 5, 400),
    ("TOR-001", "Torta de chocolate (mediana)", "Tortas y postres", 120.0, 4, 1, 2),
]


def columnas_sqlite(db, tabla):
    return [fila[1] for fila in db.execute(f"PRAGMA table_info({tabla})")]


def migrar_sqlite_a_negocios(db):
    """Bases locales anteriores (un solo negocio): todo pasa al negocio 1 sin perder datos."""
    if not columnas_sqlite(db, "usuarios") or "negocio_id" in columnas_sqlite(db, "usuarios"):
        return
    if "demo" not in columnas_sqlite(db, "ventas"):
        db.execute("ALTER TABLE ventas ADD COLUMN demo INTEGER NOT NULL DEFAULT 0")
        db.commit()
    db.execute("PRAGMA foreign_keys = OFF")
    nombre = db.execute("SELECT valor FROM config WHERE clave = 'negocio_nombre'").fetchone()
    nombre = nombre[0] if nombre else CONFIG_INICIAL["negocio_nombre"]
    tablas = ESQUEMA.split("CREATE TABLE IF NOT EXISTS ")
    crear = {t.split(" ", 1)[0]: "CREATE TABLE " + t.strip().rstrip(";") for t in tablas[1:]}
    db.executescript(f"""
        BEGIN;
        {crear['negocios'].split(';')[0]};
        INSERT INTO negocios (id, nombre) VALUES (1, '{nombre.replace("'", "''")}');
        ALTER TABLE usuarios ADD COLUMN negocio_id INTEGER NOT NULL DEFAULT 1 REFERENCES negocios(id);
        ALTER TABLE ventas ADD COLUMN negocio_id INTEGER NOT NULL DEFAULT 1 REFERENCES negocios(id);
        {crear['config'].split(';')[0].replace('CREATE TABLE config', 'CREATE TABLE config_nueva')};
        INSERT INTO config_nueva (negocio_id, clave, valor) SELECT 1, clave, valor FROM config;
        DROP TABLE config;
        ALTER TABLE config_nueva RENAME TO config;
        {crear['productos'].split(';')[0].replace('CREATE TABLE productos', 'CREATE TABLE productos_nueva')};
        INSERT INTO productos_nueva (id, negocio_id, codigo, nombre, categoria, precio, stock,
                                     stock_minimo, fecha_vencimiento, activo)
            SELECT id, 1, codigo, nombre, categoria, precio, stock, stock_minimo,
                   fecha_vencimiento, activo FROM productos;
        DROP TABLE productos;
        ALTER TABLE productos_nueva RENAME TO productos;
        COMMIT;
    """)
    db.execute("PRAGMA foreign_keys = ON")
    problemas = db.execute("PRAGMA foreign_key_check").fetchall()
    if problemas:
        raise RuntimeError(f"Migración con referencias rotas: {problemas[:5]}")


def poblar_negocio(db, negocio_id, nombre, ejemplos=True):
    for clave, valor in CONFIG_INICIAL.items():
        if clave == "negocio_nombre":
            valor = nombre
        db.execute("""INSERT INTO config (negocio_id, clave, valor) VALUES (?, ?, ?)
                      ON CONFLICT (negocio_id, clave) DO NOTHING""", (negocio_id, clave, valor))
    if ejemplos and not db.execute("SELECT COUNT(*) FROM productos WHERE negocio_id = ?",
                                   (negocio_id,)).fetchone()[0]:
        base = ahora().date()
        for cod, nom, cat, precio, stock, minimo, dias in PRODUCTOS_EJEMPLO:
            venc = (base + timedelta(days=dias)).isoformat() if dias is not None else None
            db.execute("""INSERT INTO productos (negocio_id, codigo, nombre, categoria, precio, stock,
                          stock_minimo, fecha_vencimiento) VALUES (?,?,?,?,?,?,?,?)""",
                       (negocio_id, cod, nom, cat, precio, stock, minimo, venc))


def crear_negocio(nombre, usuario, clave, ejemplos=True):
    """Crea un negocio con su cuenta principal (rol ventas). Devuelve el id del negocio."""
    usuario = usuario.strip().lower()
    if len(clave) < 8:
        raise ValueError("La contraseña debe tener al menos 8 caracteres.")
    db = conectar()
    try:
        bloquear_escrituras(db)
        if db.execute("SELECT 1 FROM usuarios WHERE usuario = ?", (usuario,)).fetchone():
            raise ValueError(f"El usuario «{usuario}» ya existe.")
        negocio_id = db.execute("INSERT INTO negocios (nombre) VALUES (?) RETURNING id",
                                (nombre,)).fetchone()[0]
        db.execute("""INSERT INTO usuarios (negocio_id, usuario, nombre, rol, clave_hash)
                      VALUES (?,?,?,?,?)""",
                   (negocio_id, usuario, nombre, "ventas", generate_password_hash(clave)))
        poblar_negocio(db, negocio_id, nombre, ejemplos)
        db.commit()
        return negocio_id
    finally:
        db.close()


def init_db():
    if ES_PG:
        db = conectar()
        db.execute("SELECT pg_advisory_xact_lock(4243)")  # evita carreras entre arranques simultáneos
        db.execute(esquema_pg())
        db.commit()
        db.close()
        return  # en la nube los negocios se crean con crear_negocio.py, sin cuentas por defecto
    db = conectar()
    migrar_sqlite_a_negocios(db)
    db.executescript(ESQUEMA)
    if db.execute("SELECT COUNT(*) FROM negocios").fetchone()[0] == 0:
        db.execute("INSERT INTO negocios (id, nombre) VALUES (1, ?)", (CONFIG_INICIAL["negocio_nombre"],))
    if db.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0] == 0:
        db.execute("""INSERT INTO usuarios (negocio_id, usuario, nombre, rol, clave_hash)
                      VALUES (1,?,?,?,?)""",
                   ("ventas", "Encargado de Ventas", "ventas", generate_password_hash("ventas123")))
        db.execute("""INSERT INTO usuarios (negocio_id, usuario, nombre, rol, clave_hash)
                      VALUES (1,?,?,?,?)""",
                   ("envios", "Encargado de Envíos", "envios", generate_password_hash("envios123")))
    poblar_negocio(db, 1, CONFIG_INICIAL["negocio_nombre"])
    db.commit()
    db.close()


def obtener_config():
    if g.get("usuario") is None:
        return {**CONFIG_INICIAL, "negocio_nombre": NOMBRE_SITIO}
    if "config" not in g:
        filas = get_db().execute("SELECT clave, valor FROM config WHERE negocio_id = ?",
                                 (nid(),)).fetchall()
        g.config = {**CONFIG_INICIAL, **{f["clave"]: f["valor"] for f in filas}}
    return g.config


# ---------------------------------------------------------------- autenticación
@app.before_request
def cargar_usuario():
    uid = session.get("usuario_id")
    g.usuario = None
    if uid:
        g.usuario = get_db().execute("SELECT * FROM usuarios WHERE id = ?", (uid,)).fetchone()


def requiere_rol(*roles):
    def decorador(vista):
        @wraps(vista)
        def envoltura(*args, **kwargs):
            if g.usuario is None:
                return redirect(url_for("login", next=request.path))
            if roles and g.usuario["rol"] not in roles:
                abort(403)
            return vista(*args, **kwargs)
        return envoltura
    return decorador


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        fila = get_db().execute("SELECT * FROM usuarios WHERE usuario = ?",
                                (request.form.get("usuario", "").strip().lower(),)).fetchone()
        if fila and check_password_hash(fila["clave_hash"], request.form.get("clave", "")):
            session.clear()
            session["usuario_id"] = fila["id"]
            destino = request.args.get("next")
            if destino and destino.startswith("/") and not destino.startswith("//"):
                return redirect(destino)
            return redirect(url_for("inicio"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/cuenta", methods=["GET", "POST"])
@requiere_rol()
def cuenta():
    if request.method == "POST":
        actual = request.form.get("actual", "")
        nueva = request.form.get("nueva", "")
        if not check_password_hash(g.usuario["clave_hash"], actual):
            flash("La contraseña actual no es correcta.", "error")
        elif len(nueva) < 6:
            flash("La nueva contraseña debe tener al menos 6 caracteres.", "error")
        elif nueva != request.form.get("confirmar", ""):
            flash("Las contraseñas no coinciden.", "error")
        else:
            db = get_db()
            db.execute("UPDATE usuarios SET clave_hash = ? WHERE id = ?",
                       (generate_password_hash(nueva), g.usuario["id"]))
            db.commit()
            flash("Contraseña actualizada.", "ok")
            return redirect(url_for("inicio"))
    return render_template("cuenta.html")


@app.errorhandler(403)
def prohibido(_e):
    return render_template("error.html", mensaje="No tenés permiso para ver esta página."), 403


@app.errorhandler(404)
def no_encontrado(_e):
    return render_template("error.html", mensaje="Página no encontrada."), 404


# ---------------------------------------------------------------- inicio
def alertas_inventario():
    db = get_db()
    dias = int(obtener_config().get("dias_aviso_vencimiento", 7))
    limite = (hoy() + timedelta(days=dias)).isoformat()
    stock_bajo = db.execute("""SELECT * FROM productos WHERE negocio_id = ? AND activo = 1
                               AND stock <= stock_minimo ORDER BY stock""", (nid(),)).fetchall()
    por_vencer = db.execute("""SELECT * FROM productos WHERE negocio_id = ? AND activo = 1 AND stock > 0
                               AND fecha_vencimiento IS NOT NULL AND fecha_vencimiento <= ?
                               ORDER BY fecha_vencimiento""", (nid(), limite)).fetchall()
    return stock_bajo, por_vencer


@app.route("/")
@requiere_rol()
def inicio():
    db = get_db()
    if g.usuario["rol"] == "envios":
        return redirect(url_for("envios"))
    dia = hoy().isoformat()
    resumen = db.execute("""SELECT COUNT(*) AS n, COALESCE(SUM(total),0) AS total FROM ventas
                            WHERE negocio_id = ? AND anulada = 0 AND substr(fecha,1,10) = ?""",
                         (nid(), dia)).fetchone()
    envios_pend = db.execute("""SELECT COUNT(*) FROM envios e JOIN ventas v ON v.id = e.venta_id
                             WHERE v.negocio_id = ? AND e.estado IN ('pendiente','en_camino')""",
                             (nid(),)).fetchone()[0]
    stock_bajo, por_vencer = alertas_inventario()
    ultimas = db.execute("""SELECT * FROM ventas WHERE negocio_id = ? ORDER BY id DESC LIMIT 8""",
                         (nid(),)).fetchall()
    return render_template("inicio.html", resumen=resumen, envios_pend=envios_pend,
                           stock_bajo=stock_bajo, por_vencer=por_vencer, ultimas=ultimas,
                           hoy=hoy().isoformat(), graficos=graficos_inicio())


DIAS_SEMANA = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def graficos_inicio():
    """Datos de los gráficos del inicio. Cada fila: etq, valor, texto (valor formateado), tip."""
    db = get_db()
    dia = hoy()
    filas = db.execute("""SELECT CAST(substr(fecha,12,2) AS INTEGER) AS h, COUNT(*) AS n,
                           SUM(total) AS total FROM ventas WHERE negocio_id = ? AND anulada = 0
                           AND substr(fecha,1,10) = ? GROUP BY h""", (nid(), dia.isoformat())).fetchall()
    por_h = {r["h"]: r for r in filas}
    h_ini = min([8] + list(por_h))
    h_fin = max([ahora().hour, h_ini] + list(por_h))
    por_hora = []
    for h in range(h_ini, h_fin + 1):
        r = por_h.get(h)
        total, n = (r["total"], r["n"]) if r else (0, 0)
        por_hora.append({"etq": f"{h}h", "valor": total, "texto": formato_bs(total),
                         "tip": f"{h}:00 a {h}:59 · {n} venta(s) · {formato_bs(total)}"})

    desde7 = (dia - timedelta(days=6)).isoformat()
    filas = db.execute("""SELECT substr(fecha,1,10) AS d, COUNT(*) AS n, SUM(total) AS total
                           FROM ventas WHERE negocio_id = ? AND anulada = 0
                           AND substr(fecha,1,10) BETWEEN ? AND ?
                           GROUP BY d""", (nid(), desde7, dia.isoformat())).fetchall()
    por_d = {r["d"]: r for r in filas}
    semana = []
    for i in range(6, -1, -1):
        d = dia - timedelta(days=i)
        r = por_d.get(d.isoformat())
        total, n = (r["total"], r["n"]) if r else (0, 0)
        etq = "Hoy" if i == 0 else f"{DIAS_SEMANA[d.weekday()]} {d.day}"
        semana.append({"etq": etq, "valor": total, "texto": formato_bs(total),
                       "tip": f"{formato_fecha(d)} · {n} venta(s) · {formato_bs(total)}"})

    cierre, _, _ = datos_cierre(dia.isoformat())
    metodos = [{"etq": f["metodo"], "valor": f["total"], "texto": formato_bs(f["total"]),
                "tip": f"{f['metodo']}: {f['n']} venta(s) · {formato_bs(f['total'])}"} for f in cierre]

    _, top, _ = datos_periodo(dia.replace(day=1).isoformat(), dia.isoformat())
    productos_top = [{"etq": r["nombre"], "valor": r["cantidad"], "texto": f"{r['cantidad']} u.",
                      "tip": f"{r['nombre']}: {r['cantidad']} u. · {formato_bs(r['total'])}"}
                     for r in top[:5]]
    return {"por_hora": por_hora, "semana": semana, "metodos": metodos, "top": productos_top,
            "hay_hoy": any(f["valor"] for f in por_hora)}


# ---------------------------------------------------------------- productos
@app.route("/productos")
@requiere_rol("ventas")
def productos():
    q = request.args.get("q", "").strip()
    sql = "SELECT * FROM productos WHERE negocio_id = ? AND activo = 1"
    params = [nid()]
    if q:
        sql += " AND (LOWER(nombre) LIKE ? OR LOWER(codigo) LIKE ? OR LOWER(categoria) LIKE ?)"
        params += [f"%{q.lower()}%"] * 3
    sql += " ORDER BY categoria, nombre"
    lista = get_db().execute(sql, params).fetchall()
    dias = int(obtener_config().get("dias_aviso_vencimiento", 7))
    return render_template("productos.html", productos=lista, q=q, hoy=hoy().isoformat(),
                           limite_venc=(hoy() + timedelta(days=dias)).isoformat())


def leer_form_producto():
    f = request.form
    errores = []
    datos = {
        "codigo": f.get("codigo", "").strip() or None,
        "nombre": f.get("nombre", "").strip(),
        "categoria": f.get("categoria", "Otros"),
        "fecha_vencimiento": f.get("fecha_vencimiento") or None,
    }
    if not datos["nombre"]:
        errores.append("El nombre es obligatorio.")
    try:
        datos["precio"] = round(float(f.get("precio", "0").replace(",", ".")), 2)
        if datos["precio"] < 0:
            raise ValueError
    except ValueError:
        errores.append("Precio inválido.")
    for campo in ("stock", "stock_minimo"):
        try:
            datos[campo] = int(f.get(campo, "0"))
            if datos[campo] < 0:
                raise ValueError
        except ValueError:
            errores.append(f"Valor inválido en {campo.replace('_', ' ')}.")
    return datos, errores


@app.route("/productos/nuevo", methods=["GET", "POST"])
@app.route("/productos/<int:pid>/editar", methods=["GET", "POST"])
@requiere_rol("ventas")
def producto_form(pid=None):
    db = get_db()
    producto = None
    if pid:
        producto = db.execute("SELECT * FROM productos WHERE id = ? AND negocio_id = ?",
                              (pid, nid())).fetchone()
        if not producto:
            abort(404)
    if request.method == "POST":
        datos, errores = leer_form_producto()
        if not errores:
            try:
                if producto:
                    diferencia = datos["stock"] - producto["stock"]
                    db.execute("""UPDATE productos SET codigo=?, nombre=?, categoria=?, precio=?,
                                  stock=?, stock_minimo=?, fecha_vencimiento=? WHERE id=? AND negocio_id=?""",
                               (datos["codigo"], datos["nombre"], datos["categoria"], datos["precio"],
                                datos["stock"], datos["stock_minimo"], datos["fecha_vencimiento"], pid,
                                nid()))
                    if diferencia:
                        registrar_movimiento(pid, diferencia, "Ajuste manual")
                else:
                    nuevo_id = db.execute("""INSERT INTO productos (negocio_id, codigo, nombre,
                                        categoria, precio, stock, stock_minimo, fecha_vencimiento)
                                        VALUES (?,?,?,?,?,?,?,?) RETURNING id""",
                                          (nid(), datos["codigo"], datos["nombre"], datos["categoria"],
                                           datos["precio"], datos["stock"], datos["stock_minimo"],
                                           datos["fecha_vencimiento"])).fetchone()[0]
                    if datos["stock"]:
                        registrar_movimiento(nuevo_id, datos["stock"], "Stock inicial")
                db.commit()
                flash("Producto guardado.", "ok")
                return redirect(url_for("productos"))
            except ErrorIntegridad:
                db.rollback()  # Postgres deja la transacción inválida tras el error
                errores.append("Ya existe un producto con ese código.")
        for e in errores:
            flash(e, "error")
        producto = {**(dict(producto) if producto else {}), **datos}
    return render_template("producto_form.html", producto=producto, categorias=CATEGORIAS)


@app.route("/productos/<int:pid>/reponer", methods=["POST"])
@requiere_rol("ventas")
def producto_reponer(pid):
    db = get_db()
    try:
        cantidad = int(request.form.get("cantidad", "0"))
    except ValueError:
        cantidad = 0
    if cantidad <= 0:
        flash("Ingresá una cantidad mayor a cero.", "error")
        return redirect(url_for("productos"))
    venc = request.form.get("fecha_vencimiento") or None
    if not db.execute("SELECT 1 FROM productos WHERE id = ? AND negocio_id = ?", (pid, nid())).fetchone():
        abort(404)
    db.execute("UPDATE productos SET stock = stock + ? WHERE id = ?", (cantidad, pid))
    if venc:
        db.execute("UPDATE productos SET fecha_vencimiento = ? WHERE id = ?", (venc, pid))
    registrar_movimiento(pid, cantidad, "Reposición")
    db.commit()
    flash("Stock repuesto.", "ok")
    return redirect(url_for("productos"))


@app.route("/productos/<int:pid>/eliminar", methods=["POST"])
@requiere_rol("ventas")
def producto_eliminar(pid):
    db = get_db()
    db.execute("UPDATE productos SET activo = 0 WHERE id = ? AND negocio_id = ?", (pid, nid()))
    db.commit()
    flash("Producto dado de baja.", "ok")
    return redirect(url_for("productos"))


def registrar_movimiento(producto_id, cantidad, motivo):
    get_db().execute("""INSERT INTO movimientos_stock (fecha, producto_id, cantidad, motivo, usuario_id)
                        VALUES (?,?,?,?,?)""",
                     (ahora().isoformat(timespec="seconds"), producto_id, cantidad, motivo,
                      g.usuario["id"] if g.usuario else None))


# ---------------------------------------------------------------- ventas
@app.route("/ventas/nueva", methods=["GET", "POST"])
@requiere_rol("ventas")
def venta_nueva():
    db = get_db()
    cfg = obtener_config()
    tarifa = float(cfg.get("tarifa_envio", 0))
    lista = db.execute("""SELECT * FROM productos WHERE negocio_id = ? AND activo = 1 AND stock > 0
                          ORDER BY nombre""", (nid(),)).fetchall()
    if request.method == "GET":
        return render_template("venta_nueva.html", productos=lista, tarifa=tarifa,
                               hoy=hoy().isoformat(), form={})

    f = request.form
    errores = []
    ids = f.getlist("producto_id")
    cants = f.getlist("cantidad")
    items = {}
    for pid, cant in zip(ids, cants):
        try:
            pid, cant = int(pid), int(cant)
        except ValueError:
            continue
        if cant > 0:
            items[pid] = items.get(pid, 0) + cant
    if not items:
        errores.append("Agregá al menos un producto.")

    detalle = []
    subtotal = 0.0
    hoy_iso = hoy().isoformat()
    for pid, cant in items.items():
        p = db.execute("SELECT * FROM productos WHERE id = ? AND negocio_id = ? AND activo = 1",
                       (pid, nid())).fetchone()
        if not p:
            errores.append("Producto no encontrado.")
            continue
        if p["stock"] < cant:
            errores.append(f"Stock insuficiente de «{p['nombre']}» (disponible: {p['stock']}).")
        if p["fecha_vencimiento"] and p["fecha_vencimiento"] < hoy_iso:
            errores.append(f"«{p['nombre']}» está vencido y no se puede vender.")
        total_item = round(p["precio"] * cant, 2)
        subtotal += total_item
        detalle.append((p, cant, total_item))

    con_envio = f.get("con_envio") == "1"
    costo_envio = tarifa if con_envio else 0.0
    total = round(subtotal + costo_envio, 2)

    metodo = f.get("metodo_pago")
    if metodo not in METODOS_PAGO:
        errores.append("Elegí un método de pago.")
    monto_recibido = cambio = None
    referencia = f.get("referencia_pago", "").strip() or None
    if metodo == "efectivo":
        try:
            monto_recibido = round(float(f.get("monto_recibido", "0").replace(",", ".")), 2)
        except ValueError:
            monto_recibido = 0
        if not math.isfinite(monto_recibido) or monto_recibido < total:
            errores.append("El monto recibido en efectivo es menor al total.")
        cambio = round(monto_recibido - total, 2)
    elif metodo in ("qr", "transferencia") and not referencia:
        errores.append("Ingresá el número de transacción / comprobante.")

    envio = None
    if con_envio:
        envio = {k: f.get(k, "").strip() for k in
                 ("destinatario", "telefono", "direccion", "referencia", "fecha_programada")}
        if not envio["destinatario"]:
            envio["destinatario"] = f.get("cliente_nombre", "").strip()
        if not (envio["destinatario"] and envio["telefono"] and envio["direccion"]):
            errores.append("Para el envío completá destinatario, teléfono y dirección.")

    if errores:
        for e in errores:
            flash(e, "error")
        return render_template("venta_nueva.html", productos=lista, tarifa=tarifa,
                               hoy=hoy_iso, form=f, items=items)

    momento = ahora().isoformat(timespec="seconds")
    venta_id = db.execute("""INSERT INTO ventas (negocio_id, fecha, usuario_id, cliente_nombre,
                        cliente_nit, subtotal, costo_envio, total, metodo_pago, monto_recibido, cambio,
                        referencia_pago) VALUES (?,?,?,?,?,?,?,?,?,?,?,?) RETURNING id""",
                          (nid(), momento, g.usuario["id"], f.get("cliente_nombre", "").strip() or "S/N",
                           f.get("cliente_nit", "").strip() or "0", round(subtotal, 2), costo_envio,
                           total, metodo, monto_recibido, cambio, referencia)).fetchone()[0]
    for p, cant, total_item in detalle:
        db.execute("""INSERT INTO venta_items (venta_id, producto_id, nombre, cantidad,
                      precio_unitario, total) VALUES (?,?,?,?,?,?)""",
                   (venta_id, p["id"], p["nombre"], cant, p["precio"], total_item))
        db.execute("UPDATE productos SET stock = stock - ? WHERE id = ?", (cant, p["id"]))
        registrar_movimiento(p["id"], -cant, f"Venta #{venta_id}")
    if envio:
        db.execute("""INSERT INTO envios (venta_id, destinatario, telefono, direccion, referencia,
                      fecha_programada, estado, actualizado) VALUES (?,?,?,?,?,?, 'pendiente', ?)""",
                   (venta_id, envio["destinatario"], envio["telefono"], envio["direccion"],
                    envio["referencia"] or None, envio["fecha_programada"] or None, momento))
    db.commit()
    flash(f"Venta #{venta_id} registrada.", "ok")
    return redirect(url_for("venta_detalle", vid=venta_id))


@app.route("/ventas")
@requiere_rol("ventas")
def ventas():
    desde = request.args.get("desde") or hoy().isoformat()
    hasta = request.args.get("hasta") or hoy().isoformat()
    lista = get_db().execute("""SELECT v.*, e.estado AS estado_envio FROM ventas v
                                LEFT JOIN envios e ON e.venta_id = v.id
                                WHERE v.negocio_id = ? AND substr(v.fecha,1,10) BETWEEN ? AND ?
                                ORDER BY v.id DESC""", (nid(), desde, hasta)).fetchall()
    return render_template("ventas.html", ventas=lista, desde=desde, hasta=hasta)


def cargar_venta(vid):
    db = get_db()
    venta = db.execute("""SELECT v.*, u.nombre AS vendedor FROM ventas v
                          JOIN usuarios u ON u.id = v.usuario_id
                          WHERE v.id = ? AND v.negocio_id = ?""", (vid, nid())).fetchone()
    if not venta:
        abort(404)
    items = db.execute("SELECT * FROM venta_items WHERE venta_id = ?", (vid,)).fetchall()
    envio = db.execute("SELECT * FROM envios WHERE venta_id = ?", (vid,)).fetchone()
    return venta, items, envio


@app.route("/ventas/<int:vid>")
@requiere_rol("ventas")
def venta_detalle(vid):
    venta, items, envio = cargar_venta(vid)
    return render_template("venta_detalle.html", venta=venta, items=items, envio=envio)


@app.route("/ventas/<int:vid>/nota")
@requiere_rol("ventas", "envios")
def venta_nota(vid):
    venta, items, envio = cargar_venta(vid)
    return render_template("nota_venta.html", venta=venta, items=items, envio=envio)


@app.route("/ventas/<int:vid>/anular", methods=["POST"])
@requiere_rol("ventas")
def venta_anular(vid):
    db = get_db()
    venta, items, envio = cargar_venta(vid)
    if venta["anulada"]:
        flash("La venta ya estaba anulada.", "error")
        return redirect(url_for("venta_detalle", vid=vid))
    if envio and envio["estado"] == "entregado":
        flash("No se puede anular una venta cuyo envío ya fue entregado.", "error")
        return redirect(url_for("venta_detalle", vid=vid))
    motivo = request.form.get("motivo", "").strip()
    if not motivo:
        flash("Indicá el motivo de la anulación.", "error")
        return redirect(url_for("venta_detalle", vid=vid))
    db.execute("UPDATE ventas SET anulada = 1, motivo_anulacion = ? WHERE id = ?", (motivo, vid))
    for it in items:
        db.execute("UPDATE productos SET stock = stock + ? WHERE id = ?", (it["cantidad"], it["producto_id"]))
        registrar_movimiento(it["producto_id"], it["cantidad"], f"Anulación venta #{vid}")
    if envio:
        db.execute("UPDATE envios SET estado = 'cancelado', actualizado = ? WHERE venta_id = ?",
                   (ahora().isoformat(timespec="seconds"), vid))
    db.commit()
    flash(f"Venta #{vid} anulada y stock devuelto.", "ok")
    return redirect(url_for("venta_detalle", vid=vid))


# ---------------------------------------------------------------- envíos
@app.route("/envios")
@requiere_rol("ventas", "envios")
def envios():
    estado = request.args.get("estado", "activos")
    sql = """SELECT e.*, v.total, v.fecha AS fecha_venta, v.cliente_nombre, u.nombre AS repartidor
             FROM envios e JOIN ventas v ON v.id = e.venta_id
             LEFT JOIN usuarios u ON u.id = e.repartidor_id WHERE v.negocio_id = ?"""
    params = [nid()]
    if estado == "activos":
        sql += " AND e.estado IN ('pendiente','en_camino')"
    elif estado in ESTADOS_ENVIO:
        sql += " AND e.estado = ?"
        params.append(estado)
    sql += " ORDER BY COALESCE(e.fecha_programada, substr(v.fecha,1,10)), e.id"
    lista = get_db().execute(sql, params).fetchall()
    return render_template("envios.html", envios=lista, estado=estado, hoy=hoy().isoformat())


SIGUIENTE_ESTADO = {"pendiente": "en_camino", "en_camino": "entregado"}


@app.route("/envios/<int:eid>/estado", methods=["POST"])
@requiere_rol("ventas", "envios")
def envio_estado(eid):
    db = get_db()
    envio = db.execute("""SELECT e.* FROM envios e JOIN ventas v ON v.id = e.venta_id
                          WHERE e.id = ? AND v.negocio_id = ?""", (eid, nid())).fetchone()
    if not envio:
        abort(404)
    nuevo = request.form.get("estado")
    permitido = SIGUIENTE_ESTADO.get(envio["estado"])
    if envio["estado"] == "en_camino" and nuevo == "pendiente":
        permitido = "pendiente"  # devolver a pendiente (no se pudo entregar)
    if nuevo != permitido:
        flash("Cambio de estado no permitido.", "error")
        return redirect(url_for("envios"))
    obs = request.form.get("observaciones", "").strip()
    db.execute("""UPDATE envios SET estado = ?, actualizado = ?, repartidor_id = ?,
                  observaciones = CASE WHEN ? <> '' THEN ? ELSE observaciones END WHERE id = ?""",
               (nuevo, ahora().isoformat(timespec="seconds"), g.usuario["id"], obs, obs, eid))
    db.commit()
    flash(f"Envío #{eid}: {ESTADOS_ENVIO[nuevo]}.", "ok")
    return redirect(request.referrer or url_for("envios"))


# ---------------------------------------------------------------- reportes
def datos_cierre(dia):
    db = get_db()
    por_metodo = db.execute("""SELECT metodo_pago, COUNT(*) AS n, SUM(total) AS total
                               FROM ventas WHERE negocio_id = ? AND anulada = 0
                               AND substr(fecha,1,10) = ? GROUP BY metodo_pago""", (nid(), dia)).fetchall()
    totales = db.execute("""SELECT COUNT(*) AS n, COALESCE(SUM(total),0) AS total,
                            COALESCE(SUM(costo_envio),0) AS envios FROM ventas
                            WHERE negocio_id = ? AND anulada = 0 AND substr(fecha,1,10) = ?""",
                         (nid(), dia)).fetchone()
    anuladas = db.execute("""SELECT COUNT(*) AS n, COALESCE(SUM(total),0) AS total FROM ventas
                             WHERE negocio_id = ? AND anulada = 1 AND substr(fecha,1,10) = ?""",
                          (nid(), dia)).fetchone()
    filas = {m: {"metodo": nombre, "n": 0, "total": 0.0} for m, nombre in METODOS_PAGO.items()}
    for r in por_metodo:
        filas[r["metodo_pago"]].update(n=r["n"], total=r["total"])
    return list(filas.values()), totales, anuladas


def datos_periodo(desde, hasta):
    db = get_db()
    por_dia = db.execute("""SELECT substr(fecha,1,10) AS dia, COUNT(*) AS n, SUM(total) AS total
                            FROM ventas WHERE negocio_id = ? AND anulada = 0
                            AND substr(fecha,1,10) BETWEEN ? AND ?
                            GROUP BY dia ORDER BY dia""", (nid(), desde, hasta)).fetchall()
    top = db.execute("""SELECT MAX(i.nombre) AS nombre, SUM(i.cantidad) AS cantidad, SUM(i.total) AS total
                        FROM venta_items i JOIN ventas v ON v.id = i.venta_id
                        WHERE v.negocio_id = ? AND v.anulada = 0 AND substr(v.fecha,1,10) BETWEEN ? AND ?
                        GROUP BY i.producto_id ORDER BY cantidad DESC, total DESC LIMIT 15""",
                     (nid(), desde, hasta)).fetchall()
    total = sum(r["total"] for r in por_dia)
    return por_dia, top, total


def datos_envios(desde, hasta):
    db = get_db()
    por_estado = db.execute("""SELECT e.estado, COUNT(*) AS n, COALESCE(SUM(v.costo_envio),0) AS cobrado
                               FROM envios e JOIN ventas v ON v.id = e.venta_id
                               WHERE v.negocio_id = ? AND substr(v.fecha,1,10) BETWEEN ? AND ?
                               GROUP BY e.estado""", (nid(), desde, hasta)).fetchall()
    detalle = db.execute("""SELECT e.*, v.fecha AS fecha_venta, v.costo_envio, v.total,
                            u.nombre AS repartidor FROM envios e JOIN ventas v ON v.id = e.venta_id
                            LEFT JOIN usuarios u ON u.id = e.repartidor_id
                            WHERE v.negocio_id = ? AND substr(v.fecha,1,10) BETWEEN ? AND ?
                            ORDER BY e.id""", (nid(), desde, hasta)).fetchall()
    filas = {k: {"estado": v, "n": 0, "cobrado": 0.0} for k, v in ESTADOS_ENVIO.items()}
    for r in por_estado:
        filas[r["estado"]].update(n=r["n"], cobrado=r["cobrado"])
    return list(filas.values()), detalle


def rango_fechas():
    hasta = request.args.get("hasta") or hoy().isoformat()
    desde = request.args.get("desde") or (hoy().replace(day=1)).isoformat()
    if desde > hasta:
        desde, hasta = hasta, desde
    return desde, hasta


@app.route("/reportes")
@requiere_rol("ventas")
def reportes():
    dia = request.args.get("dia") or hoy().isoformat()
    desde, hasta = rango_fechas()
    cierre, totales, anuladas = datos_cierre(dia)
    por_dia, top, total_periodo = datos_periodo(desde, hasta)
    env_estado, env_detalle = datos_envios(desde, hasta)
    return render_template("reportes.html", dia=dia, desde=desde, hasta=hasta, cierre=cierre,
                           totales=totales, anuladas=anuladas, por_dia=por_dia, top=top,
                           total_periodo=total_periodo, env_estado=env_estado,
                           env_detalle=env_detalle, grafico_dias=[
                               {"etq": r["dia"][8:10] + "/" + r["dia"][5:7], "valor": r["total"],
                                "texto": formato_bs(r["total"]),
                                "tip": f"{formato_fecha(r['dia'])} · {r['n']} venta(s) · {formato_bs(r['total'])}"}
                               for r in por_dia])


def hoja(wb, titulo, encabezados, filas, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = titulo
    ws.append(encabezados)
    for c in ws[1]:
        c.font = Font(bold=True)
    for fila in filas:
        ws.append(list(fila))
    for col in ws.columns:
        ancho = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(ancho + 2, 10), 50)
    return ws


@app.route("/reportes/excel/<tipo>")
@requiere_rol("ventas")
def reporte_excel(tipo):
    wb = Workbook()
    if tipo == "cierre":
        dia = request.args.get("dia") or hoy().isoformat()
        cierre, totales, anuladas = datos_cierre(dia)
        filas = [(c["metodo"], c["n"], round(c["total"], 2)) for c in cierre]
        filas += [("TOTAL", totales["n"], round(totales["total"], 2)),
                  ("Incluye envíos cobrados", "", round(totales["envios"], 2)),
                  ("Ventas anuladas (no suman)", anuladas["n"], round(anuladas["total"], 2))]
        hoja(wb, "Cierre de caja", ["Método de pago", "N° ventas", "Total Bs"], filas, primera=True)
        ventas_dia = get_db().execute("""SELECT id, fecha, cliente_nombre, cliente_nit, metodo_pago,
                                         referencia_pago, total, anulada FROM ventas
                                         WHERE negocio_id = ? AND substr(fecha,1,10) = ?
                                         ORDER BY id""", (nid(), dia)).fetchall()
        hoja(wb, "Ventas del día", ["N°", "Fecha", "Cliente", "NIT/CI", "Método", "Referencia",
                                    "Total Bs", "Anulada"],
             [(v["id"], v["fecha"][:16].replace("T", " "), v["cliente_nombre"], v["cliente_nit"],
               METODOS_PAGO.get(v["metodo_pago"]), v["referencia_pago"] or "", v["total"],
               "Sí" if v["anulada"] else "No") for v in ventas_dia])
        nombre = f"cierre_caja_{dia}.xlsx"
    elif tipo == "periodo":
        desde, hasta = rango_fechas()
        por_dia, top, total = datos_periodo(desde, hasta)
        hoja(wb, "Ventas por día", ["Fecha", "N° ventas", "Total Bs"],
             [(r["dia"], r["n"], round(r["total"], 2)) for r in por_dia] + [("TOTAL", "", round(total, 2))],
             primera=True)
        hoja(wb, "Más vendidos", ["Producto", "Cantidad", "Total Bs"],
             [(r["nombre"], r["cantidad"], round(r["total"], 2)) for r in top])
        nombre = f"ventas_{desde}_a_{hasta}.xlsx"
    elif tipo == "envios":
        desde, hasta = rango_fechas()
        por_estado, detalle = datos_envios(desde, hasta)
        hoja(wb, "Resumen envíos", ["Estado", "Cantidad", "Cobrado por envío Bs"],
             [(r["estado"], r["n"], round(r["cobrado"], 2)) for r in por_estado], primera=True)
        hoja(wb, "Detalle envíos", ["N° envío", "Venta", "Fecha venta", "Destinatario", "Teléfono",
                                    "Dirección", "Estado", "Costo envío Bs", "Repartidor", "Actualizado"],
             [(e["id"], e["venta_id"], e["fecha_venta"][:16].replace("T", " "), e["destinatario"],
               e["telefono"], e["direccion"], ESTADOS_ENVIO.get(e["estado"]), e["costo_envio"],
               e["repartidor"] or "", (e["actualizado"] or "")[:16].replace("T", " ")) for e in detalle])
        nombre = f"envios_{desde}_a_{hasta}.xlsx"
    elif tipo == "inventario":
        prods = get_db().execute("""SELECT * FROM productos WHERE negocio_id = ? AND activo = 1
                                    ORDER BY categoria, nombre""", (nid(),)).fetchall()
        hoja(wb, "Inventario", ["Código", "Producto", "Categoría", "Precio Bs", "Stock", "Stock mínimo",
                                "Vencimiento", "Valor Bs"],
             [(p["codigo"] or "", p["nombre"], p["categoria"], p["precio"], p["stock"], p["stock_minimo"],
               p["fecha_vencimiento"] or "", round(p["precio"] * p["stock"], 2)) for p in prods],
             primera=True)
        nombre = f"inventario_{hoy().isoformat()}.xlsx"
    else:
        abort(404)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(buffer, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------- configuración
@app.route("/configuracion", methods=["GET", "POST"])
@requiere_rol("ventas")
def configuracion():
    db = get_db()
    if request.method == "POST":
        errores = []
        try:
            tarifa = float(request.form.get("tarifa_envio", "0").replace(",", "."))
            if tarifa < 0:
                raise ValueError
        except ValueError:
            errores.append("Tarifa de envío inválida.")
        try:
            dias = int(request.form.get("dias_aviso_vencimiento", "7"))
            if dias < 0:
                raise ValueError
        except ValueError:
            errores.append("Días de aviso inválidos.")
        if errores:
            for e in errores:
                flash(e, "error")
        else:
            for clave in CONFIG_INICIAL:
                valor = request.form.get(clave, "").strip()
                if clave == "tarifa_envio":
                    valor = f"{tarifa:.2f}"
                db.execute("""INSERT INTO config (negocio_id, clave, valor) VALUES (?, ?, ?)
                              ON CONFLICT (negocio_id, clave) DO UPDATE SET valor = excluded.valor""",
                           (nid(), clave, valor))
            db.commit()
            flash("Configuración guardada.", "ok")
            return redirect(url_for("configuracion"))
    n_demo = get_db().execute("SELECT COUNT(*) FROM ventas WHERE negocio_id = ? AND demo = 1",
                              (nid(),)).fetchone()[0]
    return render_template("configuracion.html", n_demo=n_demo)


# ---------------------------------------------------------------- datos de prueba (demo)
MOTIVO_REPOSICION_DEMO = "Reposición (datos de prueba)"
CLIENTES_DEMO = [
    ("María Quispe", "4567890"), ("Panadería El Trigal", "1023456019"),
    ("Café La Glorieta", "3021457012"), ("Juan Mamani", "6789012"), ("S/N", "0"), ("S/N", "0"),
    ("Repostería Dulce Sucre", "2034567011"), ("Ana Flores", "5432109"), ("Carlos Vargas", "7654321"),
]
DIRECCIONES_DEMO = [
    "Calle Junín 245, zona Central", "Av. Hernando Siles 890, zona Mercado Campesino",
    "Calle Bolívar 532, zona Central", "Av. Las Américas 1200, zona San Matías",
    "Calle Arenales 77, zona Surapata",
]


def respaldar_antes_de_demo():
    """Copia la base una sola vez, antes de generar datos de prueba por primera vez."""
    if ES_PG:
        return  # en la nube no hay archivo local; las ventas de prueba se pueden borrar después
    destino = os.path.join(BASE_DIR, "reposteria.antes-de-datos-de-prueba.db")
    if not os.path.exists(destino):
        copia = sqlite3.connect(destino)
        get_db().backup(copia)
        copia.close()


def generar_ventas_demo(cantidad=12):
    db = get_db()
    hoy_iso = hoy().isoformat()
    vigente = "(fecha_vencimiento IS NULL OR fecha_vencimiento >= ?)"
    # Reponer lo que está bajo para que haya algo que vender.
    dias_aviso = int(obtener_config().get("dias_aviso_vencimiento", 7))
    lejos = (hoy() + timedelta(days=dias_aviso)).isoformat()
    for p in db.execute("""SELECT id FROM productos WHERE negocio_id = ? AND activo = 1 AND stock < 15
                           AND (fecha_vencimiento IS NULL OR fecha_vencimiento > ?)""",
                        (nid(), lejos)).fetchall():
        db.execute("UPDATE productos SET stock = stock + 20 WHERE id = ?", (p["id"],))
        registrar_movimiento(p["id"], 20, MOTIVO_REPOSICION_DEMO)
    vendibles = db.execute(f"""SELECT * FROM productos WHERE negocio_id = ? AND activo = 1
                               AND stock > 0 AND {vigente}""", (nid(), hoy_iso)).fetchall()
    if not vendibles:
        return 0
    stock = {p["id"]: p["stock"] for p in vendibles}
    tarifa = float(obtener_config().get("tarifa_envio", 0))
    repartidor = db.execute("SELECT id FROM usuarios WHERE negocio_id = ? AND rol = 'envios' LIMIT 1",
                            (nid(),)).fetchone()

    fin = ahora()
    inicio = fin.replace(hour=8, minute=0, second=0, microsecond=0)
    if fin - inicio < timedelta(minutes=30):
        inicio = fin.replace(hour=0, minute=0, second=0, microsecond=0)
    segundos = max((fin - inicio).total_seconds(), 60)
    momentos = sorted(inicio + timedelta(seconds=random.uniform(0, segundos)) for _ in range(cantidad))

    creadas = []
    for momento_dt in momentos:
        momento = momento_dt.isoformat(timespec="seconds")
        detalle = []
        for p in random.sample(vendibles, k=min(len(vendibles), random.randint(1, 4))):
            if stock[p["id"]] <= 0:
                continue
            cant = random.randint(1, min(3, stock[p["id"]]))
            stock[p["id"]] -= cant
            detalle.append((p, cant, round(p["precio"] * cant, 2)))
        if not detalle:
            continue
        subtotal = round(sum(d[2] for d in detalle), 2)
        con_envio = random.random() < 0.3
        costo_envio = tarifa if con_envio else 0.0
        total = round(subtotal + costo_envio, 2)
        metodo = random.choices(list(METODOS_PAGO), weights=[45, 35, 10, 10])[0]
        recibido = cambio = referencia = None
        if metodo == "efectivo":
            recibido = float(math.ceil(total / 10) * 10 + random.choice([0, 0, 10, 50]))
            cambio = round(recibido - total, 2)
        elif metodo == "qr":
            referencia = f"QR-{random.randint(10**7, 10**8 - 1)}"
        elif metodo == "transferencia":
            referencia = f"TRF-{random.randint(100000, 999999)}"
        else:
            referencia = f"VCH-{random.randint(1000, 9999)}"
        cliente, nit = random.choice(CLIENTES_DEMO)
        vid = db.execute("""INSERT INTO ventas (negocio_id, fecha, usuario_id, cliente_nombre,
                            cliente_nit, subtotal, costo_envio, total, metodo_pago, monto_recibido,
                            cambio, referencia_pago, demo) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,1) RETURNING id""",
                         (nid(), momento, g.usuario["id"], cliente, nit, subtotal, costo_envio, total,
                          metodo, recibido, cambio, referencia)).fetchone()[0]
        for p, cant, total_item in detalle:
            db.execute("""INSERT INTO venta_items (venta_id, producto_id, nombre, cantidad,
                          precio_unitario, total) VALUES (?,?,?,?,?,?)""",
                       (vid, p["id"], p["nombre"], cant, p["precio"], total_item))
            db.execute("UPDATE productos SET stock = stock - ? WHERE id = ?", (cant, p["id"]))
            db.execute("""INSERT INTO movimientos_stock (fecha, producto_id, cantidad, motivo, usuario_id)
                          VALUES (?,?,?,?,?)""", (momento, p["id"], -cant, f"Venta #{vid}", g.usuario["id"]))
        if con_envio:
            estado = random.choices(["pendiente", "en_camino", "entregado"], weights=[40, 30, 30])[0]
            destinatario = cliente if cliente != "S/N" else "Cliente de prueba"
            db.execute("""INSERT INTO envios (venta_id, destinatario, telefono, direccion, fecha_programada,
                          estado, repartidor_id, actualizado) VALUES (?,?,?,?,?,?,?,?)""",
                       (vid, destinatario, f"7{random.randint(1000000, 9999999)}",
                        random.choice(DIRECCIONES_DEMO), hoy_iso, estado,
                        repartidor["id"] if repartidor and estado != "pendiente" else None, momento))
        creadas.append((vid, con_envio, detalle))

    # Una venta anulada, para ver cómo se muestra.
    sin_envio = [c for c in creadas if not c[1]]
    if len(creadas) >= 5 and sin_envio:
        vid, _, detalle = random.choice(sin_envio)
        db.execute("UPDATE ventas SET anulada = 1, motivo_anulacion = ? WHERE id = ?",
                   ("Cliente se arrepintió (prueba)", vid))
        for p, cant, _ in detalle:
            db.execute("UPDATE productos SET stock = stock + ? WHERE id = ?", (cant, p["id"]))
            registrar_movimiento(p["id"], cant, f"Anulación venta #{vid}")
    db.commit()
    return len(creadas)


def borrar_datos_demo():
    db = get_db()
    ventas_demo = db.execute("SELECT id, anulada FROM ventas WHERE negocio_id = ? AND demo = 1",
                             (nid(),)).fetchall()
    for v in ventas_demo:
        if not v["anulada"]:
            for it in db.execute("SELECT producto_id, cantidad FROM venta_items WHERE venta_id = ?",
                                 (v["id"],)).fetchall():
                db.execute("UPDATE productos SET stock = stock + ? WHERE id = ?",
                           (it["cantidad"], it["producto_id"]))
        db.execute("DELETE FROM movimientos_stock WHERE motivo IN (?, ?)",
                   (f"Venta #{v['id']}", f"Anulación venta #{v['id']}"))
        db.execute("DELETE FROM envios WHERE venta_id = ?", (v["id"],))
        db.execute("DELETE FROM venta_items WHERE venta_id = ?", (v["id"],))
        db.execute("DELETE FROM ventas WHERE id = ?", (v["id"],))
    for r in db.execute("""SELECT m.producto_id, SUM(m.cantidad) AS c FROM movimientos_stock m
                           JOIN productos p ON p.id = m.producto_id
                           WHERE m.motivo = ? AND p.negocio_id = ? GROUP BY m.producto_id""",
                        (MOTIVO_REPOSICION_DEMO, nid())).fetchall():
        db.execute("""UPDATE productos SET stock = CASE WHEN stock > ? THEN stock - ? ELSE 0 END
                      WHERE id = ?""", (r["c"], r["c"], r["producto_id"]))
    db.execute("""DELETE FROM movimientos_stock WHERE motivo = ?
                  AND producto_id IN (SELECT id FROM productos WHERE negocio_id = ?)""",
               (MOTIVO_REPOSICION_DEMO, nid()))
    db.commit()
    return len(ventas_demo)


@app.route("/demo/generar", methods=["POST"])
@requiere_rol("ventas")
def demo_generar():
    respaldar_antes_de_demo()
    n = generar_ventas_demo()
    flash(f"Se crearon {n} ventas de prueba de hoy (con pagos, envíos y una anulación).", "ok")
    return redirect(url_for("inicio"))


@app.route("/demo/borrar", methods=["POST"])
@requiere_rol("ventas")
def demo_borrar():
    n = borrar_datos_demo()
    flash(f"Se borraron {n} ventas de prueba y se devolvió el stock.", "ok")
    return redirect(request.referrer or url_for("inicio"))


# ---------------------------------------------------------------- asistente (chatbot por reglas)
import unicodedata

from flask import jsonify

PALABRAS_VACIAS = {
    "a", "al", "cual", "cuanto", "cuanta", "cuantos", "cuantas", "cuesta", "cuestan", "de",
    "del", "el", "en", "es", "esta", "hay", "la", "las", "lo", "los", "me", "mi", "precio",
    "precios", "que", "queda", "quedan", "se", "stock", "tenemos", "tiene", "un", "una",
    "unidades", "vale", "valen", "y", "dime", "decime", "sale", "salen", "producto", "por",
    "favor", "porfa", "cuantas", "existe", "existen", "busca", "buscar", "info", "sobre",
}

AYUDA_VENTAS = ("Puedo responder cosas como:\n"
                "1. ¿Cuánto vendimos hoy?\n"
                "2. ¿Qué productos tienen stock bajo?\n"
                "3. ¿Qué está vencido o por vencer?\n"
                "4. ¿Cuántos envíos hay pendientes?\n"
                "5. ¿Qué es lo más vendido este mes?\n"
                "6. ¿Cuál es la tarifa de envío?\n"
                "7. Precio o stock de un producto, por ejemplo: «precio fondant»")
AYUDA_ENVIOS = ("Puedo responder cosas como:\n"
                "1. ¿Cuántos envíos hay pendientes?\n"
                "2. ¿Qué envíos están en camino?\n"
                "3. ¿Cuál es la tarifa de envío?")


def normalizar(texto):
    texto = unicodedata.normalize("NFD", texto.lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9ñ\- ]+", " ", texto)


def contiene(texto, *claves):
    return any(re.search(rf"\b{c}", texto) for c in claves)


def bot_ventas_hoy():
    filas, totales, anuladas = datos_cierre(hoy().isoformat())
    if not totales["n"]:
        return "Hoy todavía no hay ventas registradas."
    lineas = [f"Hoy van {totales['n']} venta(s) por {formato_bs(totales['total'])}."]
    numero = 1
    for f in filas:
        if f["n"]:
            lineas.append(f"{numero}. {f['metodo']}: {f['n']} · {formato_bs(f['total'])}")
            numero += 1
    if anuladas["n"]:
        lineas.append(f"Anuladas: {anuladas['n']} ({formato_bs(anuladas['total'])}).")
    return "\n".join(lineas)


def bot_stock_bajo():
    stock_bajo, _ = alertas_inventario()
    if not stock_bajo:
        return "Ningún producto está en stock bajo. 👍"
    lineas = [f"{len(stock_bajo)} producto(s) con stock bajo:"]
    for i, p in enumerate(stock_bajo[:10], 1):
        lineas.append(f"{i}. {p['nombre']}: {p['stock']} (mínimo {p['stock_minimo']})")
    if len(stock_bajo) > 10:
        lineas.append(f"…y {len(stock_bajo) - 10} más en Productos.")
    return "\n".join(lineas)


def bot_vencimientos():
    _, por_vencer = alertas_inventario()
    if not por_vencer:
        return "No hay productos vencidos ni por vencer."
    dia = hoy().isoformat()
    lineas = ["Productos vencidos o por vencer:"]
    for i, p in enumerate(por_vencer[:10], 1):
        estado = "VENCIDO" if p["fecha_vencimiento"] < dia else "vence"
        lineas.append(f"{i}. {p['nombre']}: {estado} el {formato_fecha(date.fromisoformat(p['fecha_vencimiento']))}"
                      f" (stock {p['stock']})")
    return "\n".join(lineas)


def bot_envios(solo_en_camino=False):
    db = get_db()
    estados = ("en_camino",) if solo_en_camino else ("pendiente", "en_camino")
    marcas = ",".join("?" * len(estados))
    lista = db.execute(f"""SELECT e.id, e.destinatario, e.direccion, e.estado, e.fecha_programada
                           FROM envios e JOIN ventas v ON v.id = e.venta_id
                           WHERE v.negocio_id = ? AND e.estado IN ({marcas})
                           ORDER BY COALESCE(e.fecha_programada, ''), e.id""",
                       (nid(), *estados)).fetchall()
    if not lista:
        return "No hay envíos en camino." if solo_en_camino else "No hay envíos pendientes. 🎉"
    titulo = "en camino" if solo_en_camino else "activos (pendientes o en camino)"
    lineas = [f"{len(lista)} envío(s) {titulo}:"]
    for i, e in enumerate(lista[:10], 1):
        lineas.append(f"{i}. #{e['id']} {e['destinatario']} · {e['direccion']} · {ESTADOS_ENVIO[e['estado']]}")
    return "\n".join(lineas)


def bot_mas_vendidos():
    desde = hoy().replace(day=1).isoformat()
    _, top, _ = datos_periodo(desde, hoy().isoformat())
    if not top:
        return "Este mes todavía no hay ventas."
    lineas = ["Lo más vendido este mes:"]
    for i, r in enumerate(top[:5], 1):
        lineas.append(f"{i}. {r['nombre']}: {r['cantidad']} u. · {formato_bs(r['total'])}")
    return "\n".join(lineas)


def bot_producto(texto):
    palabras = [p for p in texto.split() if p not in PALABRAS_VACIAS and len(p) > 1]
    if not palabras:
        return None
    productos_activos = get_db().execute(
        "SELECT * FROM productos WHERE negocio_id = ? AND activo = 1 ORDER BY nombre", (nid(),)).fetchall()
    encontrados = [p for p in productos_activos
                   if all(w in normalizar(f"{p['nombre']} {p['codigo'] or ''} {p['categoria']}")
                          for w in palabras)]
    if not encontrados:
        return None
    lineas = []
    for i, p in enumerate(encontrados[:8], 1):
        lineas.append(f"{i}. {p['nombre']}: {formato_bs(p['precio'])} · stock {p['stock']}")
    if len(encontrados) > 8:
        lineas.append(f"…y {len(encontrados) - 8} más. Probá con un nombre más específico.")
    return "\n".join(lineas)


def responder_chat(mensaje):
    texto = normalizar(mensaje).strip()
    es_ventas = g.usuario["rol"] == "ventas"
    ayuda = AYUDA_VENTAS if es_ventas else AYUDA_ENVIOS
    if not texto:
        return ayuda
    if contiene(texto, "hola", "buen", "ayuda", "que puedes", "que podes", "opciones", "menu"):
        return f"¡Hola, {g.usuario['nombre']}! {ayuda}"
    if contiene(texto, "gracias"):
        return "¡De nada! 🧁"
    if contiene(texto, "tarifa", "costo de envio", "cuanto cobra"):
        return f"La tarifa de envío es {formato_bs(obtener_config().get('tarifa_envio', 0))}."
    if contiene(texto, "envio", "pedido", "domicilio", "entrega", "reparto"):
        return bot_envios(solo_en_camino=contiene(texto, "camino"))
    if not es_ventas:
        return "Solo puedo ayudarte con los envíos. " + ayuda
    if contiene(texto, "vend") and contiene(texto, "mas vendid", "top", "mejor", "popular"):
        return bot_mas_vendidos()
    if contiene(texto, "vend", "venta", "caja", "cierre", "factur", "ingreso") and \
            contiene(texto, "hoy", "dia", "caja", "cierre", "cuanto"):
        return bot_ventas_hoy()
    if contiene(texto, "stock bajo", "poco stock", "reponer", "falta", "agot", "quedan poco"):
        return bot_stock_bajo()
    if contiene(texto, "venc", "caduc"):
        return bot_vencimientos()
    respuesta = bot_producto(texto)
    if respuesta:
        return respuesta
    return "No entendí la pregunta ni encontré un producto con ese nombre. " + ayuda


@app.route("/asistente", methods=["POST"])
@requiere_rol()
def asistente():
    datos = request.get_json(silent=True) or {}
    mensaje = str(datos.get("mensaje", ""))[:300]
    return jsonify(respuesta=responder_chat(mensaje))


TABLAS_RESPALDO = {  # solo las filas del negocio que descarga el respaldo
    "negocios": "id = ?",
    "usuarios": "negocio_id = ?",
    "config": "negocio_id = ?",
    "productos": "negocio_id = ?",
    "ventas": "negocio_id = ?",
    "venta_items": "venta_id IN (SELECT id FROM ventas WHERE negocio_id = ?)",
    "envios": "venta_id IN (SELECT id FROM ventas WHERE negocio_id = ?)",
    "movimientos_stock": "producto_id IN (SELECT id FROM productos WHERE negocio_id = ?)",
}


def exportar_negocio_sqlite(ruta):
    """Copia los datos del negocio desde Postgres a un archivo SQLite con el mismo esquema."""
    db = get_db()
    destino = sqlite3.connect(ruta)
    destino.executescript(ESQUEMA)
    for tabla, condicion in TABLAS_RESPALDO.items():
        filas = db.execute(f"SELECT * FROM {tabla} WHERE {condicion} ORDER BY 1", (nid(),)).fetchall()
        if filas:
            columnas = list(filas[0].keys())
            destino.executemany(
                f"INSERT INTO {tabla} ({', '.join(columnas)}) VALUES ({', '.join('?' * len(columnas))})",
                [tuple(f[c] for c in columnas) for f in filas])
    destino.commit()
    destino.close()


@app.route("/respaldo", methods=["POST"])
@requiere_rol("ventas")
def descargar_respaldo():
    with tempfile.TemporaryDirectory() as carpeta:
        ruta = os.path.join(carpeta, "respaldo.db")
        if ES_PG:
            exportar_negocio_sqlite(ruta)
        else:
            # Conexión independiente: backup consistente incluso con el servidor en marcha.
            get_db().rollback()
            with sqlite3.connect(DB_PATH) as origen, sqlite3.connect(ruta) as destino:
                origen.backup(destino)
        with open(ruta, "rb") as archivo:
            contenido = io.BytesIO(archivo.read())
    return send_file(contenido, as_attachment=True,
                     download_name=f"reposteria-{ahora():%Y%m%d-%H%M%S}.db",
                     mimetype="application/octet-stream")


init_db()

if __name__ == "__main__":
    from waitress import serve
    serve(app, host=os.environ.get("HOST", "0.0.0.0"),
          port=int(os.environ.get("PORT", "5000")), threads=4)
