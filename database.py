
import sqlite3
import json
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")

os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, "auto_form.db")


def connect():

    return sqlite3.connect(DB_PATH)


def create_database():

    conn = connect()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS profiles (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT NOT NULL,
            identity_number TEXT,
            birth_date TEXT,
            gender TEXT,
            address TEXT,
            phone TEXT,
            occupation TEXT,
            extra_data TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    conn.commit()
    conn.close()
    create_template_tables()


def save_profile(data, profile_id=None):

    conn = connect()
    cursor = conn.cursor()

    extra_data = {
        key: value
        for key, value in data.items()
        if key not in [
            "HO_TEN",
            "CCCD",
            "NGAY_SINH",
            "GIOI_TINH",
            "DIA_CHI",
            "SO_DIEN_THOAI",
            "NGHE_NGHIEP"
        ]
    }

    values = (
        data.get("HO_TEN", ""),
        data.get("CCCD", ""),
        data.get("NGAY_SINH", ""),
        data.get("GIOI_TINH", ""),
        data.get("DIA_CHI", ""),
        data.get("SO_DIEN_THOAI", ""),
        data.get("NGHE_NGHIEP", ""),
        json.dumps(extra_data, ensure_ascii=False)
    )

    if profile_id:

        cursor.execute("""
            UPDATE profiles SET
                full_name=?,
                identity_number=?,
                birth_date=?,
                gender=?,
                address=?,
                phone=?,
                occupation=?,
                extra_data=?,
                updated_at=CURRENT_TIMESTAMP
            WHERE id=?
        """, values + (profile_id,))

    else:

        cursor.execute("""
            INSERT INTO profiles (
                full_name,
                identity_number,
                birth_date,
                gender,
                address,
                phone,
                occupation,
                extra_data
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, values)

        profile_id = cursor.lastrowid

    conn.commit()
    conn.close()

    return profile_id


def get_profiles(keyword=""):

    conn = connect()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, full_name, identity_number, birth_date
        FROM profiles
        WHERE full_name LIKE ?
           OR identity_number LIKE ?
        ORDER BY updated_at DESC
    """, (
        f"%{keyword}%",
        f"%{keyword}%"
    ))

    result = cursor.fetchall()

    conn.close()

    return result


def get_profile(profile_id):

    conn = connect()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT * FROM profiles WHERE id=?
    """, (profile_id,))

    row = cursor.fetchone()

    conn.close()

    if not row:
        return None

    return {
        "HO_TEN": row[1],
        "CCCD": row[2],
        "NGAY_SINH": row[3],
        "GIOI_TINH": row[4],
        "DIA_CHI": row[5],
        "SO_DIEN_THOAI": row[6],
        "NGHE_NGHIEP": row[7],
        **json.loads(row[8] or "{}")
    }


def delete_profile(profile_id):

    conn = connect()
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM profiles WHERE id=?",
        (profile_id,)
    )

    conn.commit()
    conn.close()


if __name__ == "__main__":

    create_database()

    print("Đã khởi tạo cơ sở dữ liệu AUTO FORM!")


# Thu vien bieu mau: chi xoa ban ghi, khong xoa tep Word.
def create_template_tables():
    with connect() as conn:
        conn.execute("CREATE TABLE IF NOT EXISTS template_categories (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE)")
        conn.execute("CREATE TABLE IF NOT EXISTS templates (id INTEGER PRIMARY KEY AUTOINCREMENT, category_id INTEGER NOT NULL, name TEXT NOT NULL, file_path TEXT NOT NULL, UNIQUE(category_id,name))")
        conn.execute("CREATE TABLE IF NOT EXISTS app_settings (setting_key TEXT PRIMARY KEY, setting_value TEXT NOT NULL)")

def get_template_categories():
    with connect() as conn:
        return conn.execute("SELECT id,name FROM template_categories ORDER BY name COLLATE NOCASE").fetchall()

def add_template_category(name):
    with connect() as conn:
        conn.execute("INSERT INTO template_categories(name) VALUES(?)", (name,))

def rename_template_category(category_id, name):
    with connect() as conn:
        conn.execute("UPDATE template_categories SET name=? WHERE id=?", (name,category_id))

def delete_template_category(category_id):
    with connect() as conn:
        conn.execute("DELETE FROM templates WHERE category_id=?", (category_id,))
        conn.execute("DELETE FROM template_categories WHERE id=?", (category_id,))

def get_templates(category_id=None):
    with connect() as conn:
        if category_id is None:
            return conn.execute("SELECT id,category_id,name,file_path FROM templates ORDER BY name COLLATE NOCASE").fetchall()
        return conn.execute("SELECT id,category_id,name,file_path FROM templates WHERE category_id=? ORDER BY name COLLATE NOCASE",(category_id,)).fetchall()

def add_template(category_id,name,file_path):
    with connect() as conn:
        conn.execute("INSERT INTO templates(category_id,name,file_path) VALUES(?,?,?)",(category_id,name,file_path))

def rename_template(template_id,name):
    with connect() as conn:
        conn.execute("UPDATE templates SET name=? WHERE id=?",(name,template_id))

def replace_template_file(template_id,path):
    with connect() as conn:
        conn.execute("UPDATE templates SET file_path=? WHERE id=?",(path,template_id))

def delete_template(template_id):
    with connect() as conn:
        conn.execute("DELETE FROM templates WHERE id=?",(template_id,))

def get_setting(key):
    with connect() as conn:
        row=conn.execute("SELECT setting_value FROM app_settings WHERE setting_key=?",(key,)).fetchone()
        return row[0] if row else None

def set_setting(key,value):
    with connect() as conn:
        conn.execute("INSERT OR REPLACE INTO app_settings(setting_key,setting_value) VALUES(?,?)",(key,value))
