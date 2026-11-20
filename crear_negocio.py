"""Da de alta un negocio nuevo con su cuenta principal.

Uso (contra la base de Vercel/Neon):
    DATABASE_URL="postgresql://..." .venv/bin/python crear_negocio.py "Repostería Carito" reposteria_carito

La contraseña se pide por teclado (o se genera si se deja vacía) y no queda guardada en ningún archivo.
Agregue --sin-ejemplos para empezar sin los productos de ejemplo.
"""
import getpass
import secrets
import sys

import app as sistema


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 2:
        sys.exit(__doc__)
    nombre, usuario = args
    clave = getpass.getpass("Contraseña (vacía = generar una): ") if sys.stdin.isatty() else ""
    clave = clave or secrets.token_urlsafe(9)
    negocio_id = sistema.crear_negocio(nombre, usuario, clave,
                                       ejemplos="--sin-ejemplos" not in sys.argv)
    print(f"Negocio #{negocio_id} «{nombre}» creado.")
    print(f"Usuario: {usuario.strip().lower()}")
    print(f"Contraseña: {clave}")


if __name__ == "__main__":
    main()
