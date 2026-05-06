"""One-shot bootstrap: create the database (if missing) and all tables.

Usage:
    python3 setup_db.py

Reads credentials from .env. Never prints the password.
"""

import os
import sys

from db import ensure_database, init_schema, health_check


def main():
    host = os.environ.get("DB_HOST", "?")
    name = os.environ.get("DB_NAME", "?")
    print(f"Target: {host} / {name}")
    print()
    print("1/3 Creating database if missing...")
    try:
        ensure_database()
    except Exception as e:
        print(f"  ERROR: {e}")
        sys.exit(1)
    print("  OK")

    print("2/3 Creating tables...")
    try:
        init_schema()
    except Exception as e:
        print(f"  ERROR: {e}")
        sys.exit(1)
    print("  OK")

    print("3/3 Health check...")
    if not health_check():
        print("  ERROR: could not query database after setup")
        sys.exit(1)
    print("  OK")
    print()
    print("Done.")


if __name__ == "__main__":
    # Load .env before db.py is imported if needed
    from dotenv import load_dotenv
    load_dotenv()
    main()
