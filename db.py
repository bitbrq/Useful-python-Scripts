import sqlite3


def connect():
    try:
        conn = sqlite3.connect('contacts.db')
        return conn
    except sqlite3.Error as e:
        print("Error connecting to database:", e)
        return None


def initialize_database():
    conn = connect()
    cursor = conn.cursor()

    # Create tables if not exists
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS birthday (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            birthday TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS name (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS full_name (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS organization (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            organization TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS address (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            address TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS number (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS email (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS photo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            photo BLOB NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS number_label (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            number_label TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS contact (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            birthday_id INTEGER,
            name_id INTEGER,
            full_name_id INTEGER,  -- New column for storing the full name ID
            organization_id INTEGER,  -- New column for storing the organization ID
            address_id INTEGER,  -- New column for storing the address ID
            number_label_id INTEGER,
            number_id INTEGER,
            email_id INTEGER,
            photo_id INTEGER,  -- New column for storing the photo ID
            FOREIGN KEY (birthday_id) REFERENCES birthday(id),
            FOREIGN KEY (name_id) REFERENCES name(id),
            FOREIGN KEY (full_name_id) REFERENCES full_name(id),
            FOREIGN KEY (organization_id) REFERENCES organization(id),
            FOREIGN KEY (address_id) REFERENCES address(id),
            FOREIGN KEY (number_label_id) REFERENCES number_label(id),
            FOREIGN KEY (number_id) REFERENCES number(id),
            FOREIGN KEY (email_id) REFERENCES email(id),
            FOREIGN KEY (photo_id) REFERENCES photo(id)
        )
    ''')

    conn.commit()
    conn.close()


def create_record(name_id, number_id, email_id):
    conn = connect()
    cursor = conn.cursor()

    cursor.execute("INSERT INTO contact (name_id, number_id, email_id) VALUES (?, ?, ?)",
                   (name_id, number_id, email_id))

    conn.commit()
    conn.close()


def get_records():
    conn = connect()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM contact")
    rows = cursor.fetchall()

    conn.close()
    return rows


def insert_record(record, table_name, column_name):
    conn = connect()
    cursor = conn.cursor()

    # Check if the record already exists
    cursor.execute(f"SELECT * FROM {table_name} WHERE {column_name} = ?", (record,))
    rows = cursor.fetchall()

    if len(rows) > 0:
        # Record already exists, fetch its ID
        cursor.execute(f"SELECT id FROM {table_name} WHERE {column_name} = ?", (record,))
        id = cursor.fetchone()[0]
        conn.close()
        return id
    else:
        # Insert the record into the table
        cursor.execute(f"INSERT INTO {table_name} ({column_name}) VALUES (?)", (record,))
        conn.commit()
        id = cursor.lastrowid
        conn.close()
        return id


def delete_record(record_id):
    conn = connect()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM contact WHERE id = ?", (record_id,))

    conn.commit()
    conn.close()


def duplicate_verify(num):
    conn = connect()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM number WHERE number = ?", (num,))
    result = cursor.fetchone()
    conn.close()
    if result:
        return result[0]  # Return the ID if the number exists
    else:
        return None  # Return None if the number doesn't exist


def insert_contact(birthday_id, name_id, full_name_id, organization_id, address_id, label_id, num, email_id):
    conn = connect()
    cursor = conn.cursor()
    # Check for nullable values and replace them with NULL in the SQL query
    cursor.execute(
        "INSERT INTO contact (birthday_id, name_id, full_name_id, organization_id, address_id, number_label_id, number_id, email_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (birthday_id or None, name_id or None, full_name_id or None, organization_id or None, address_id or None, label_id or None,
         num or None, email_id or None))
    conn.commit()
    conn.close()
