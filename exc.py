import pandas as pd
from sqlalchemy import create_engine, text

# Fungsi untuk mengimpor data dari file Excel ke MySQL dengan kolom tambahan
def import_excel_to_db(file_path, db_url, table_name):
    # Membaca file Excel menggunakan pandas
    df = pd.read_excel(file_path)

    # Menambahkan kolom baru 'nama_user' dan 'detail' dengan nilai NULL
    df['nama_user'] = None  # Kolom nama_user akan diisi dengan NULL
    df['detail'] = None  # Kolom detail akan diisi dengan NULL

    # Membuat koneksi ke database menggunakan SQLAlchemy
    engine = create_engine(db_url)

    # Membuat tabel dengan tambahan kolom 'id' sebagai auto increment
    with engine.connect() as conn:
        conn.execute(text(f"""
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INT AUTO_INCREMENT PRIMARY KEY,
                order_no VARCHAR(255),
                idcard VARCHAR(255),
                phone VARCHAR(255),
                name VARCHAR(255),
                ocr_area TEXT,
                ocr_province VARCHAR(255),
                ocr_city VARCHAR(255),
                overdue_day INT,
                area TEXT,
                gps TEXT,
                due_date DATE,
                application_amount DECIMAL(15, 2),
                contactable TINYINT,
                mission_id INT,
                emergs_name0 VARCHAR(255),
                emergs_phone0 VARCHAR(255),
                emergs_relation0 VARCHAR(255),
                emergs_name1 VARCHAR(255),
                emergs_phone1 VARCHAR(255),
                emergs_relation1 VARCHAR(255),
                face_photo_url TEXT,
                outstanding DECIMAL(15, 2),
                repay_principal_amt DECIMAL(15, 2)
            );
        """))

    # Mengimpor data ke dalam tabel MySQL
    df.to_sql(table_name, con=engine, if_exists='append', index=False)

    print(f"Data berhasil diimpor ke tabel {table_name}")

# Lokasi file Excel
file_path = 'static/uploads/1.xlsx'

# URL koneksi ke MySQL
db_url = 'mysql+pymysql://root:@localhost/app_excel3'

# Nama tabel di database MySQL
table_name = 'data_excel'

# Menjalankan fungsi impor
import_excel_to_db(file_path, db_url, table_name)
