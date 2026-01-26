import getpass
from passlib.context import CryptContext
from db_auth import ensure_users_table, create_user

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

def main():
    ensure_users_table()
    username = input("Username: ").strip()
    password = getpass.getpass("Password: ").strip()
    password2 = getpass.getpass("Repeat Password: ").strip()

    if password != password2:
        print("Passwords do not match.")
        return

    password_hash = pwd_context.hash(password)
    create_user(username, password_hash)
    print("User created OK")

if __name__ == "__main__":
    main()
