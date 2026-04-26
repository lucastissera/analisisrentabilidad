"""
Crea un usuario en la base local. Los altas por WhatsApp se cargan con este script (o vía API en el futuro).

Uso:  python create_user.py NOMBRE_USUARIO
      (pide la contraseña de forma interactiva)

O:    python create_user.py NOMBRE_USUARIO contraseña
"""

import sys

from auth_users import create_user, init_db


def main():
    init_db()
    if len(sys.argv) < 2:
        print("Uso: python create_user.py USUARIO [contraseña]", file=sys.stderr)
        sys.exit(1)
    user = sys.argv[1]
    if len(sys.argv) >= 3:
        pwd = sys.argv[2]
    else:
        import getpass

        pwd = getpass.getpass("Contraseña: ")
        p2 = getpass.getpass("Repetir: ")
        if pwd != p2:
            print("Las contraseñas no coinciden.", file=sys.stderr)
            sys.exit(1)
    ok, err = create_user(user, pwd)
    if not ok:
        print(err, file=sys.stderr)
        sys.exit(1)
    print(f"Usuario «{user}» creado.")


if __name__ == "__main__":
    main()
