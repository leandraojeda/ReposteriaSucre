import os
import sqlite3
import tempfile
import unittest
from concurrent.futures import ThreadPoolExecutor

# Sin DATABASE_URL se prueba con SQLite temporal. Con TEST_DATABASE_URL se prueba contra un
# Postgres de pruebas, que se BORRA por completo al empezar (nunca usar la base real).
TEMP = tempfile.TemporaryDirectory()
os.environ['REPOSTERIA_DB'] = os.path.join(TEMP.name, 'test.db')
os.environ['SECRET_KEY'] = 'test-only'
os.environ.pop('DATABASE_URL', None)
os.environ.pop('POSTGRES_URL', None)
if os.environ.get('TEST_DATABASE_URL'):
    import psycopg
    os.environ['DATABASE_URL'] = os.environ['TEST_DATABASE_URL']
    with psycopg.connect(os.environ['DATABASE_URL'], autocommit=True) as pg:
        pg.execute('DROP SCHEMA public CASCADE')
        pg.execute('CREATE SCHEMA public')
import app as sistema
from werkzeug.security import generate_password_hash

if sistema.ES_PG:  # en la nube no hay cuentas por defecto: se crean como en producción
    _nid = sistema.crear_negocio('Negocio de prueba', 'ventas', 'ventas123')
    _db = sistema.conectar()
    _db.execute("INSERT INTO usuarios (negocio_id, usuario, nombre, rol, clave_hash) VALUES (?,?,?,?,?)",
                (_nid, 'envios', 'Envíos', 'envios', generate_password_hash('envios123')))
    _db.commit()
    _db.close()


def valor(sql, params=()):
    db = sistema.conectar()
    try:
        fila = db.execute(sql, params).fetchone()
        return fila[0] if fila else None
    finally:
        db.close()


def ejecutar(sql, params=()):
    db = sistema.conectar()
    try:
        db.execute(sql, params)
        db.commit()
    finally:
        db.close()


class SistemaTest(unittest.TestCase):
    def cliente(self, usuario='ventas', clave=None):
        cliente = sistema.app.test_client()
        cliente.get('/login')
        with cliente.session_transaction() as sesion:
            token = sesion['csrf_token']
        respuesta = cliente.post('/login', data=dict(usuario=usuario, clave=clave or usuario+'123', csrf_token=token))
        self.assertEqual(respuesta.status_code, 302)
        cliente.get('/', follow_redirects=True)
        return cliente

    def post(self, cliente, ruta, datos=None):
        with cliente.session_transaction() as sesion:
            token = sesion['csrf_token']
        return cliente.post(ruta, data={**(datos or {}), 'csrf_token': token})

    def vender(self, cliente, **extra):
        datos = dict(producto_id='1', cantidad='2', metodo_pago='efectivo', monto_recibido='100', **extra)
        respuesta = self.post(cliente, '/ventas/nueva', datos)
        self.assertEqual(respuesta.status_code, 302)
        return respuesta.headers['Location']

    def test_permisos_y_paginas(self):
        ventas = self.cliente()
        for ruta in ['/', '/productos', '/productos?q=harina', '/ventas', '/ventas/nueva', '/envios',
                     '/reportes', '/configuracion', '/cuenta']:
            self.assertEqual(ventas.get(ruta).status_code, 200, ruta)
        for tipo in ['cierre', 'periodo', 'envios', 'inventario']:
            self.assertEqual(ventas.get(f'/reportes/excel/{tipo}').status_code, 200, tipo)
        envios = self.cliente('envios')
        for ruta in ['/productos', '/reportes', '/configuracion', '/ventas/nueva']:
            self.assertEqual(envios.get(ruta).status_code, 403, ruta)
        self.assertEqual(self.post(envios, '/respaldo').status_code, 403)
        self.assertEqual(ventas.post('/ventas/nueva').status_code, 400)

    def test_venta_envio_anulacion_respaldo(self):
        ventas, envios = self.cliente(), self.cliente('envios')
        stock = valor('SELECT stock FROM productos WHERE id=1')
        ruta = self.vender(ventas, con_envio='1', destinatario='Cliente', telefono='70000000',
                           direccion='Sucre, centro')
        vid = int(ruta.rsplit('/', 1)[1])
        self.assertEqual(ventas.get(ruta).status_code, 200)
        self.assertEqual(ventas.get(ruta+'/nota').status_code, 200)
        self.assertEqual(valor('SELECT stock FROM productos WHERE id=1'), stock-2)
        eid = valor('SELECT id FROM envios WHERE venta_id=?', (vid,))
        self.post(envios, f'/envios/{eid}/estado', dict(estado='en_camino'))
        self.post(ventas, ruta+'/anular', dict(motivo='Prueba'))
        self.post(ventas, ruta+'/anular', dict(motivo='Repetida'))
        self.assertEqual(valor('SELECT stock FROM productos WHERE id=1'), stock)
        self.assertEqual(valor('SELECT estado FROM envios WHERE id=?', (eid,)), 'cancelado')
        respaldo = self.post(ventas, '/respaldo')
        self.assertEqual(respaldo.status_code, 200)
        ruta_backup = os.path.join(TEMP.name, 'backup.db')
        with open(ruta_backup, 'wb') as archivo:
            archivo.write(respaldo.data)
        with sqlite3.connect(ruta_backup) as db:
            self.assertEqual(db.execute('PRAGMA integrity_check').fetchone()[0], 'ok')
            self.assertEqual(db.execute('SELECT anulada FROM ventas WHERE id=?', (vid,)).fetchone()[0], 1)

    def test_datos_de_prueba_y_configuracion(self):
        ventas = self.cliente()
        self.assertEqual(self.post(ventas, '/demo/generar').status_code, 302)
        self.assertGreater(valor('SELECT COUNT(*) FROM ventas WHERE demo=1'), 0)
        self.assertEqual(ventas.get('/').status_code, 200)
        self.assertEqual(self.post(ventas, '/demo/borrar').status_code, 302)
        self.assertEqual(valor('SELECT COUNT(*) FROM ventas WHERE demo=1'), 0)
        config = dict(negocio_nombre='Mi Repostería', negocio_nit='1', negocio_direccion='Sucre',
                      negocio_telefono='', tarifa_envio='20', dias_aviso_vencimiento='5')
        self.assertEqual(self.post(ventas, '/configuracion', config).status_code, 302)
        self.assertIn('Mi Repostería', ventas.get('/').get_data(as_text=True))

    def test_negocios_aislados(self):
        sistema.crear_negocio('Repostería Aislada', 'aislada_prueba', 'clave-aislada', ejemplos=False)
        propio = self.cliente()
        ajeno = self.cliente('aislada_prueba', 'clave-aislada')
        ruta = self.vender(propio)
        vid = int(ruta.rsplit('/', 1)[1])
        stock = valor('SELECT stock FROM productos WHERE id=1')
        inicio = ajeno.get('/').get_data(as_text=True)
        self.assertIn('Repostería Aislada', inicio)
        self.assertNotIn('Harina 0000', ajeno.get('/productos').get_data(as_text=True))
        self.assertNotIn(f'#{vid}', ajeno.get('/ventas').get_data(as_text=True))
        for r in [ruta, ruta + '/nota', '/productos/1/editar']:
            self.assertEqual(ajeno.get(r).status_code, 404, r)
        self.assertEqual(self.post(ajeno, '/productos/1/reponer', dict(cantidad='5')).status_code, 404)
        self.assertEqual(self.post(ajeno, ruta + '/anular', dict(motivo='x')).status_code, 404)
        self.post(ajeno, '/productos/1/eliminar')
        self.assertEqual(valor('SELECT activo FROM productos WHERE id=1'), 1)
        self.assertEqual(valor('SELECT stock FROM productos WHERE id=1'), stock)
        # Un producto con el mismo código puede existir en otro negocio.
        datos = dict(codigo='HAR-001', nombre='Harina propia', categoria='Otros', precio='10',
                     stock='3', stock_minimo='1')
        self.assertEqual(self.post(ajeno, '/productos/nuevo', datos).status_code, 302)
        self.assertEqual(self.post(ajeno, '/productos/nuevo', datos).status_code, 200)  # repetido
        if sistema.ES_PG:
            respaldo = self.post(ajeno, '/respaldo')
            ruta_backup = os.path.join(TEMP.name, 'ajeno.db')
            with open(ruta_backup, 'wb') as archivo:
                archivo.write(respaldo.data)
            with sqlite3.connect(ruta_backup) as db:
                self.assertEqual(db.execute('SELECT COUNT(*) FROM ventas').fetchone()[0], 0)
                self.assertEqual(db.execute('SELECT nombre FROM productos').fetchall(), [('Harina propia',)])

    def test_ventas_simultaneas_no_sobrevenden(self):
        clientes = [self.cliente(), self.cliente()]
        ejecutar('UPDATE productos SET stock=1 WHERE id=2')
        def vender(cliente):
            return self.post(cliente, '/ventas/nueva', dict(producto_id='2', cantidad='1', metodo_pago='efectivo', monto_recibido='100')).status_code
        with ThreadPoolExecutor(max_workers=2) as pool:
            resultados = list(pool.map(vender, clientes))
        self.assertEqual(sorted(resultados), [200, 302])
        self.assertEqual(valor('SELECT stock FROM productos WHERE id=2'), 0)
