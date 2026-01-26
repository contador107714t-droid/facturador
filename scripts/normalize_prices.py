import os
import shutil
import sqlite3
from datetime import datetime
from decimal import Decimal, InvalidOperation


BASE_DIR = os.path.dirname(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "facturador.db")
BACKUP_PATH = os.path.join(BASE_DIR, "facturador_backup.db")
THRESHOLD = Decimal("10000000")


def to_decimal(value):
    if value is None:
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def normalize_value(value):
    dec = to_decimal(value)
    if dec is None:
        return value, False
    if dec == 0:
        return dec, False
    if dec % 100 == 0 and dec >= THRESHOLD:
        return dec / 100, True
    return dec, False


def backup_db():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(DB_PATH)
    if not os.path.exists(BACKUP_PATH):
        shutil.copy2(DB_PATH, BACKUP_PATH)
        return BACKUP_PATH
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    backup = os.path.join(BASE_DIR, f"facturador_backup_{stamp}.db")
    shutil.copy2(DB_PATH, backup)
    return backup


def main():
    backup = backup_db()
    print(f"Backup creado: {backup}")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(items)")
    cols = {row["name"] for row in cur.fetchall()}
    if "costo" not in cols or "precio" not in cols:
        print("Tabla items sin columnas costo/precio. Nada que normalizar.")
        return

    cur.execute("SELECT id, costo, precio FROM items")
    rows = cur.fetchall()
    updated = 0

    for row in rows:
        costo, costo_changed = normalize_value(row["costo"])
        precio, precio_changed = normalize_value(row["precio"])
        if costo_changed or precio_changed:
            cur.execute(
                "UPDATE items SET costo = ?, precio = ? WHERE id = ?",
                (float(costo), float(precio), row["id"]),
            )
            updated += 1
            print(
                f"Item {row['id']}: costo {row['costo']} -> {costo}, "
                f"precio {row['precio']} -> {precio}"
            )

    conn.commit()
    conn.close()
    print(f"Items actualizados: {updated}")


if __name__ == "__main__":
    main()
