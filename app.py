from flask import Flask, request, render_template, redirect, url_for, flash
from flask import session
from flask import send_file, abort
from flask_sqlalchemy import SQLAlchemy
from flask_bcrypt import Bcrypt
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
import pandas as pd
from sqlalchemy import create_engine, text
import io
import os
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash
import base64
from flask import Response
from geopy.geocoders import Nominatim
from datetime import datetime
import datetime
from PIL import Image, ImageDraw, ImageFont
import requests

app = Flask(__name__)
app.config['SECRET_KEY'] = 'b35dfe6ce150230940bd145823034485'
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:@localhost/app_excel3'
app.config['MAX_CONTENT_LENGTH'] = 150 * 1024 * 1024  # 150 MB
app.config['UPLOAD_FOLDER'] = 'static/uploads'

db = SQLAlchemy(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'
geolocator = Nominatim(user_agent="geoapiExercises")

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif'}  # Add other allowed extensions as needed

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Define models
class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.Enum('admin', 'user'), nullable=False)
    name = db.Column(db.String(100), nullable=True)
    staff_id = db.Column(db.String(50), unique=True, nullable=True)
    phone = db.Column(db.String(20), nullable=True)
    lokasi = db.Column(db.String(100), nullable=True)
    email = db.Column(db.String(100), unique=True, nullable=True)

class DataExcel(db.Model):
    __tablename__ = 'data_excel'

    id = db.Column(db.Integer, primary_key=True)
    order_no = db.Column(db.String(255))
    idcard = db.Column(db.String(255))
    phone = db.Column(db.String(255))
    name = db.Column(db.String(255))
    ocr_area = db.Column(db.String(255))
    ocr_province = db.Column(db.String(255))
    ocr_city = db.Column(db.String(255))
    overdue_day = db.Column(db.Integer)
    area = db.Column(db.String(255))
    gps = db.Column(db.String(500))
    due_date = db.Column(db.String(255))
    application_amount = db.Column(db.Float)
    contactable = db.Column(db.String(255))
    mission_id = db.Column(db.String(255))
    emergs_name0 = db.Column(db.String(255))
    emergs_phone0 = db.Column(db.String(255))
    emergs_relation0 = db.Column(db.String(255))
    emergs_name1 = db.Column(db.String(255))
    emergs_phone1 = db.Column(db.String(255))
    emergs_relation1 = db.Column(db.String(255))
    face_photo_url = db.Column(db.String(255))
    outstanding = db.Column(db.Float)
    repay_principal_amt = db.Column(db.Float)
    nama_user = db.Column(db.String(255))
    detail = db.Column(db.String(255))
    status = db.Column(db.String(255))
    gambar = db.Column(db.String(255)) 
    tanggal = db.Column(db.DateTime)
    user_id = db.Column(db.String(255))
    tanggal_perubahan = db.Column(db.DateTime)

@login_manager.user_loader
def load_user(user_id):
    return db.session.get(User, int(user_id))

@app.route('/')
def home():
    return render_template('login.html')

@app.route('/rules')
def rules():
    return render_template('rules_agent.html')

@app.route('/admin/statistik-user', methods=['GET'])
@login_required
def statistik_user():
    if current_user.role != 'admin':
        return redirect(url_for('index'))  # Pastikan hanya admin yang dapat mengakses

    # Pagination setup
    page = request.args.get('page', 1, type=int)
    per_page = 10  # Jumlah entri per halaman
    offset = (page - 1) * per_page

    # Ambil data semua user yang melakukan perubahan dengan pagination
    user_changes_query = db.session.query(
        DataExcel.nama_user,
        DataExcel.tanggal,
        DataExcel.gambar,
        DataExcel.detail,
        DataExcel.status
    ).join(User, User.id == DataExcel.user_id)

    # Urutkan berdasarkan tanggal terbaru (DESC)
    user_changes_query = user_changes_query.order_by(DataExcel.tanggal.desc())

    total_changes = user_changes_query.count()  # Total entri untuk pagination
    user_changes = user_changes_query.offset(offset).limit(per_page).all()  # Ambil entri untuk halaman saat ini

    total_pages = (total_changes + per_page - 1) // per_page  # Hitung total halaman

    # Hitung start_page dan end_page untuk pagination
    start_page = max(1, page - 1)  # Halaman awal untuk ditampilkan
    end_page = min(total_pages, page + 1)  # Halaman akhir untuk ditampilkan

    return render_template('statistik_user.html', user_changes=user_changes, page=page, total_pages=total_pages, start_page=start_page, end_page=end_page)

@app.route('/statistik')
@login_required
def statistik():
    from datetime import datetime

    # Ambil tanggal hari ini
    today = datetime.now().date()

    # Hitung berapa kali user yang login melakukan perubahan hari ini
    count_changes = DataExcel.query.filter(
        DataExcel.user_id == current_user.id,  # Pastikan Anda memiliki user_id di tabel
        db.func.date(DataExcel.tanggal) == today
    ).count()

    return render_template('statistik.html', target=count_changes)


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and bcrypt.check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for('dashboard'))
        flash('Login failed. Check your username and/or password.')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('You have been logged out successfully.')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.role == 'admin':
        # Ambil hanya user dengan role 'user'
        users = User.query.filter_by(role='user').all()
        return render_template('admin_dashboard.html', users=users)
    elif current_user.role == 'user':
        return redirect(url_for('show_data'))

@app.route('/edit-user/<int:user_id>', methods=['GET', 'POST'])
@login_required
def edit_user(user_id):
    user = User.query.get_or_404(user_id)
    if request.method == 'POST':
        user.username = request.form['username']
        user.staff_id = request.form['staff_id']
        user.phone = request.form['phone']
        user.lokasi = request.form['lokasi']
        user.email = request.form['email']
        
        # Cek jika password diisi
        password = request.form['password']
        if password:  # Jika ada password baru
            user.password_hash = bcrypt.generate_password_hash(password).decode('utf-8')
        
        db.session.commit()
        flash('User has been updated successfully!', 'success')
        return redirect(url_for('dashboard'))  # Adjust redirect as needed
    return render_template('edit_user.html', user=user)

@app.route('/delete-user/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    user = User.query.get_or_404(user_id)
    db.session.delete(user)
    db.session.commit()
    flash('User has been deleted successfully!', 'success')
    return redirect(url_for('dashboard'))  # Adjust redirect as needed


@app.route('/upload-excel', methods=['GET', 'POST'])
@login_required
def upload_excel():
    if current_user.role != 'admin':
        return redirect(url_for('upload_excel'))

    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file uploaded')
            return redirect(url_for('upload_excel'))

        file = request.files['file']
        if file.filename == '':
            flash('No file selected')
            return redirect(url_for('upload_excel'))

        # Read the file directly from the request object
        file_stream = io.BytesIO(file.read())
        df = pd.read_excel(file_stream, engine='openpyxl', dtype=str)

        # Nama tabel yang ada
        table_name = 'data_excel'

        # Connect to database
        engine = create_engine(app.config['SQLALCHEMY_DATABASE_URI'])
        
        # Save data to database
        try:
            df.to_sql(table_name, con=engine, if_exists='append', index=False)
            flash('File uploaded and data added successfully!')
        except Exception as e:
            flash(f'Error uploading file: {str(e)}', 'danger')

        return redirect(url_for('upload_excel'))

    # Render upload form if method is GET
    return render_template('excel.html')

@app.route('/run-query', methods=['POST'])
@login_required
def run_query():
    try:
        # Menjalankan query ALTER TABLE
        query = """
        ALTER TABLE data_excel 
        ADD COLUMN detail TEXT, 
        ADD COLUMN status VARCHAR(50), 
        ADD COLUMN gambar TEXT,
        ADD COLUMN id INT AUTO_INCREMENT PRIMARY KEY FIRST,
        ADD COLUMN tanggal DATETIME,
        ADD COLUMN user_id INT,
        ADD COLUMN tanggal_perubahan DATETIME;
        """
        
        # Eksekusi query di database
        db.session.execute(text(query))
        db.session.commit()
        
        # Menampilkan pesan sukses
        flash('Query berhasil dijalankan!', 'success')
    except Exception as e:
        # Menampilkan pesan error jika query gagal
        flash(f'Error saat menjalankan query: {str(e)}', 'danger')

    return redirect(url_for('upload_excel'))

@app.route('/show-data')
@login_required
def show_data():
    if current_user.role != 'user':
        return redirect(url_for('dashboard'))

    # Ambil status filter dari request args
    status_filter = request.args.get('status')
    search = request.args.get('search', '')  # Ambil query pencarian

    # Ambil 'name' dari pengguna yang sedang login
    username = current_user.username
    user_info = User.query.filter_by(username=username).first()
    user_name = user_info.name if user_info else None

    page = request.args.get('page', 1, type=int)
    offset = (page - 1) * 25

    # Bangun query untuk mengambil data berdasarkan 'nama_user'
    base_query = 'SELECT * FROM data_excel WHERE nama_user = %s'
    params = [user_name]

    # Tambahkan kondisi untuk mengecualikan status 'Lunas', 'Tenor', dan 'Cicilan'
    base_query += ' AND (status NOT IN (%s, %s, %s) OR status IS NULL)'
    params.extend(['Lunas', 'Tenor', 'Cicilan'])

    # Jika status filter diterapkan
    if status_filter == "NULL":
        base_query += ' AND status IS NULL'  # Menambahkan kondisi untuk status NULL
    elif status_filter:
        base_query += ' AND status = %s'
        params.append(status_filter)

    # Menambahkan filter berdasarkan search query (name atau order_no)
    if search:
        base_query += ' AND (order_no LIKE %s OR name LIKE %s)'
        params.extend([f'%{search}%', f'%{search}%'])

    # Tambahkan pengurutan untuk menempatkan status NULL terlebih dahulu
    base_query += ' ORDER BY status IS NULL, status'  # NULL muncul di depan

    # Tambahkan limit dan offset
    paginated_query = base_query + ' LIMIT %s OFFSET %s'
    params.append(25)  # Limit
    params.append(offset)  # Offset

    # Ambil data dari database
    data = pd.read_sql(paginated_query, con=db.engine, params=tuple(params))

    # Total data query
    total_data_query = 'SELECT COUNT(*) as count FROM data_excel WHERE nama_user = %s'
    total_params = [user_name]

    # Tambahkan kondisi untuk mengecualikan status 'Lunas', 'Tenor', dan 'Cicilan'
    total_data_query += ' AND (status NOT IN (%s, %s, %s) OR status IS NULL)'
    total_params.extend(['Lunas', 'Tenor', 'Cicilan'])

    # Jika status filter diterapkan
    if status_filter == "NULL":
        total_data_query += ' AND status IS NULL'
    elif status_filter:
        total_data_query += ' AND status = %s'
        total_params.append(status_filter)

    # Menambahkan filter untuk menghitung total data berdasarkan search query
    if search:
        total_data_query += ' AND (order_no LIKE %s OR name LIKE %s)'
        total_params.extend([f'%{search}%', f'%{search}%'])

    # Hitung total data
    total_data = pd.read_sql(total_data_query, con=db.engine, params=tuple(total_params))['count'][0]

    total_pages = (total_data // 25) + (1 if total_data % 25 > 0 else 0)

    # Kirim data sebagai list of dictionaries
    data_dict = data.to_dict(orient='records')

    return render_template('user_dashboard.html', data=data_dict, page=page, total_pages=total_pages, status_filter=status_filter, search=search)

@app.route('/detail-user/<order_no>')
@login_required
def detail_user(order_no):
    # Ambil detail dari data_excel berdasarkan order_no
    detail_data = DataExcel.query.filter_by(order_no=order_no).first()
    
    if detail_data is None:
        return "Data not found", 404

    return render_template('detail_user.html', detail_data=detail_data)

@app.template_filter('rupiah')
def rupiah_format(value):
    try:
        # Pastikan nilai dikonversi ke integer terlebih dahulu
        value = int(value)
        return f"Rp. {value:,.0f}".replace(',', '.')
    except (ValueError, TypeError):
        # Jika nilai tidak bisa dikonversi ke angka, kembalikan nilai aslinya
        return value


@app.route('/view-order/<int:order_id>', methods=['GET'])
@login_required
def view_order(order_id):
    # Ambil data berdasarkan order_id dari database
    query = 'SELECT * FROM data_excel WHERE order_no = :order_id'
    sql_query = text(query)
    order_data = pd.read_sql(sql_query, con=db.engine, params={'order_id': order_id})

    if order_data.empty:
        return "Order not found", 404

    return render_template('data_paid.html', order=order_data.to_dict(orient='records')[0])

@app.route('/get-image/<int:order_no>/<int:image_index>')
def get_image(order_no, image_index):
    detail_data = DataExcel.query.filter_by(order_no=order_no).first()
    
    if detail_data and detail_data.gambar:
        # Ambil list gambar dari database
        gambar_list = detail_data.gambar.split(',')
        
        if 0 <= image_index < len(gambar_list):
            image_data_base64 = gambar_list[image_index]
            image_data = base64.b64decode(image_data_base64)
            return Response(image_data, mimetype='image/png')  # Sesuaikan MIME type jika JPEG
        else:
            return "Image index out of range", 404
    else:
        return "No images found", 404

import datetime  # Pastikan import datetime

@app.route('/upload-report/<order_no>', methods=['POST'])
@login_required
def upload_report(order_no):
    detail_data = DataExcel.query.filter_by(order_no=order_no).first()
    
    if detail_data is None:
        return "Data not found", 404

    detail = request.form.get('detail')
    status = request.form.get('status')
    
    # Check and replace if fields already have values
    detail_data.detail = detail if detail else detail_data.detail
    detail_data.status = status if status else detail_data.status

    # Simpan ID pengguna yang sedang login
    detail_data.user_id = current_user.id  # Pastikan ini ada dan sesuai dengan kolom di DB

    # Save current timestamp to 'tanggal' field
    detail_data.tanggal = datetime.datetime.now()  # Atur field 'tanggal'

    # Get location details
    location, latitude, longitude = get_location()

    # Process image uploads
    gambar_files = request.files.getlist('gambar')
    if gambar_files:
        gambar_list = []
        for file in gambar_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                
                # Add watermark with location details
                if location and latitude and longitude:
                    add_watermark(file_path, location, latitude, longitude)
                    
                gambar_list.append(filename)

        # Replace old images with new ones
        if gambar_list:
            detail_data.gambar = ','.join(gambar_list)  # Replace existing images

    # Commit all changes to the database
    db.session.commit()  # Pastikan untuk commit perubahan ke database
    flash('Laporan berhasil diupload dan gambar telah diupdate!', 'success')
    return redirect(url_for('detail_user', order_no=order_no))


def get_location_details(latitude, longitude):
    geolocator = Nominatim(user_agent="excel-app")
    location = geolocator.reverse((latitude, longitude), language='id')

    if location:
        address = location.raw.get('address', {})
        kecamatan = address.get('suburb', 'Tidak Diketahui')  # Mendapatkan nama kecamatan
        kabupaten = address.get('city', 'Tidak Diketahui')  # Mendapatkan nama kabupaten/kota

        return {
            'kecamatan': kecamatan,
            'kabupaten': kabupaten,
            'latitude': latitude,
            'longitude': longitude
        }
    else:
        return None

def get_location():
    # Menggunakan layanan geolocation yang gratis
    try:
        response = requests.get('https://ipinfo.io/json')
        data = response.json()
        latitude, longitude = data['loc'].split(',')
        location_details = get_location_details(latitude, longitude)

        if location_details:
            kecamatan = location_details['kecamatan']
            kabupaten = location_details['kabupaten']
            return f"{kecamatan}, {kabupaten}", latitude, longitude
        else:
            return None, None, None
    except Exception as e:
        print("Error fetching location:", e)
        return None, None, None

def add_watermark(image_path, location, latitude, longitude):
    # Membuka gambar
    with Image.open(image_path) as img:
        draw = ImageDraw.Draw(img)

        # Mendapatkan ukuran gambar
        width, height = img.size

        # Membuat teks watermark dengan baris terpisah
        watermark_text = (
            f"Lokasi: {location}\n"
            f"Lat: {latitude}\n"
            f"Lon: {longitude}\n"
            f"Waktu: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )

        # Menggunakan font TTF dengan ukuran lebih besar
        font_path = "static/font/Raleway-Bold.ttf"  # Ganti dengan path ke font TTF Anda
        font_size = 100  # Ukuran font yang lebih besar
        font = ImageFont.truetype(font_path, font_size)

        # Mendapatkan ukuran teks untuk penentuan posisi
        text_bbox = draw.textbbox((0, 0), watermark_text, font=font)
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]

        # Menentukan posisi di kiri bawah
        x = 10  # Margin kiri
        y = height - text_height - 10  # Margin bawah

        # Menambahkan watermark ke gambar dengan pemisahan baris
        draw.multiline_text((x, y), watermark_text, fill=(0, 0, 0), font=font, spacing=2)

        # Menyimpan gambar dengan watermark
        img.save(image_path)  # Atau simpan ke lokasi lain jika perlu

from sqlalchemy import cast, Date

@app.route('/show-data-status', methods=['GET', 'POST'])
@login_required
def show_data_status():
    status = request.args.get('status', None)
    agen = request.args.get('agen', None)  # Ambil agen dari request
    tanggal = request.args.get('tanggal', None)
    search = request.args.get('search', '')  # Ambil query pencarian

    page = request.args.get('page', 1, type=int)
    per_page = 10

    # Mengambil semua agen dari tabel user yang tidak kosong
    agents = User.query.with_entities(User.name).filter(User.name != None).all()  # Ambil nama agen yang tidak None

    # Memulai query
    query = DataExcel.query.filter(DataExcel.status.in_(["Lunas", "Cicilan", "Tenor"]))

    # Filter berdasarkan status
    if status:
        query = query.filter(DataExcel.status == status)

    # Filter berdasarkan agen jika agen dipilih
    if agen:
        query = query.filter(DataExcel.nama_user == agen)

    # Filter berdasarkan tanggal jika tanggal dipilih
    if tanggal:
        query = query.filter(DataExcel.tanggal >= tanggal + " 00:00:00", DataExcel.tanggal <= tanggal + " 23:59:59")

    # Menambahkan filter berdasarkan search query (name atau order_no)
    if search:
        query = query.filter((DataExcel.name.like(f'%{search}%')) | (DataExcel.order_no.like(f'%{search}%')))

    pagination = query.paginate(page=page, per_page=per_page, error_out=False)

    total_pages = pagination.pages
    start_page = max(1, page - 2)
    end_page = min(total_pages, page + 2)

    return render_template('data_status.html', 
                           data=pagination.items, 
                           page=page, 
                           total_pages=total_pages, 
                           start_page=start_page, 
                           end_page=end_page, 
                           selected_status=status,
                           selected_agen=agen,
                           agents=agents,
                           search=search)  # Kirim nilai pencarian ke template

from flask import send_file
import pandas as pd
from io import BytesIO

@app.route('/export-excel', methods=['GET'])
@login_required
def export_excel():
    status = request.args.get('status')

    # Query untuk mengambil data berdasarkan status
    if status:
        query = DataExcel.query.filter(DataExcel.status == status).all()
    else:
        return "No status selected.", 400

    # Buat DataFrame dari data yang diambil
    data = [{'order_no': item.order_no, 
             'idcard': item.idcard, 
             'phone': item.phone, 
             'name': item.name,
             'ocr_area': item.ocr_area,
             'ocr_province': item.ocr_province,
             'ocr_city': item.ocr_city,
             'overdue_day': item.overdue_day,
             'area': item.area,
             'gps': item.gps,
             'due_date': item.due_date,
             'application_amount': item.application_amount,
             'contactable': item.contactable,
             'mission_id': item.mission_id,
             'emergs_name0': item.emergs_name0,
             'emergs_phone0': item.emergs_phone0,
             'emergs_relation0': item.emergs_relation0,
             'emergs_name1': item.emergs_name1,
             'emergs_phone1': item.emergs_phone1,
             'emergs_relation1': item.emergs_relation1,
             'face_photo_url': item.face_photo_url,
             'outstanding': item.outstanding,
             'repay_principal_amt': item.repay_principal_amt,
             'nama_collector': item.nama_user,
             'status': item.status
             }
            for item in query]
    df = pd.DataFrame(data)

    # Ekspor DataFrame ke Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')

    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f'data_{status}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/view-data')
def view_data():
    username = session.get('username')  # Ambil username dari session
    data = DataExcel.query.filter_by(nama_user=username).all()  # Query untuk mengambil data
    return render_template('excel.html', data=data)  # Kirim data ke template

@app.route('/register-user', methods=['GET', 'POST'])
@login_required
def register_user():
    if current_user.role != 'admin':
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        # Ambil semua data dari form
        username = request.form.get('username')
        password = request.form.get('password')
        name = request.form.get('name')
        staff_id = request.form.get('staff_id')
        phone = request.form.get('phone')
        email = request.form.get('email')

        # Ambil lokasi sebagai list
        lokasi = request.form.getlist('lokasi')  # Mengambil semua lokasi yang dipilih

        # Debugging: Lihat semua data yang diterima
        print("Data yang diterima:")
        print("Username:", username)
        print("Password:", password)
        print("Name:", name)
        print("Staff ID:", staff_id)
        print("Phone:", phone)
        print("Email:", email)
        print("Lokasi:", lokasi)  # Cek nilai lokasi

        # Cek nilai lokasi
        if not lokasi:  # Jika tidak ada lokasi yang dipilih
            flash('Please select at least one location.')
            return render_template('admin_dashboard.html')

        # Gabungkan lokasi menjadi string jika diperlukan
        lokasi_str = ', '.join(lokasi)  # Ubah list lokasi menjadi string dengan pemisah koma

        # Check if username or email already exists
        if User.query.filter_by(username=username).first() or User.query.filter_by(email=email).first():
            flash('Username or Email already exists.')
            return render_template('admin_dashboard.html')

        # Create new user
        hashed_password = bcrypt.generate_password_hash(password).decode('utf-8')
        new_user = User(
            username=username,
            password_hash=hashed_password,
            role='user',
            name=name,
            staff_id=staff_id,
            phone=phone,
            lokasi=lokasi_str,  # Pastikan ini sesuai dengan struktur User Anda
            email=email
        )
        db.session.add(new_user)
        db.session.commit()

        flash('User registered successfully!')
        return redirect(url_for('dashboard'))

    return render_template('admin_dashboard.html')


@app.route('/view-excel-data')
@login_required
def view_excel_data():
    page = request.args.get('page', 1, type=int)
    location = request.args.get('location', '')

    per_page = 25  # Jumlah item per halaman
    offset = (page - 1) * per_page

    # Menyusun query SQL dengan filter
    query = 'SELECT * FROM data_excel'
    filters = []
    filter_values = {}

    if location:
        filters.append('(ocr_area LIKE :loc OR ocr_province LIKE :loc OR ocr_city LIKE :loc OR area LIKE :loc)')
        filter_values['loc'] = f'%{location}%'
    
    if filters:
        query += ' WHERE ' + ' AND '.join(filters)
    
    query += ' LIMIT :per_page OFFSET :offset'
    filter_values.update({'per_page': per_page, 'offset': offset})

    # Eksekusi query untuk mendapatkan data
    sql_query = text(query)
    data = pd.read_sql(sql_query, con=db.engine, params=filter_values)

    # Total data dan halaman
    total_query = 'SELECT COUNT(*) as count FROM data_excel'
    if filters:
        total_query += ' WHERE ' + ' AND '.join(filters)
    
    total_sql_query = text(total_query)
    total_data = pd.read_sql(total_sql_query, con=db.engine, params=filter_values)['count'][0]
    total_pages = (total_data // per_page) + (1 if total_data % per_page > 0 else 0)

    # Menentukan range halaman yang ingin ditampilkan (1-10 halaman)
    start_page = max(1, page - 5)
    end_page = min(total_pages, page + 4)

    return render_template(
        'bagi_excel.html',
        data=data.to_dict(orient='records'),  # Mengirimkan data ke template
        page=page,
        total_pages=total_pages,
        start_page=start_page,
        end_page=end_page
    )

@app.route('/view-excel-data-orang')
@login_required
def view_excel_data_orang():
    page = request.args.get('page', 1, type=int)
    location = request.args.get('location', '')

    per_page = 25  # Jumlah item per halaman
    offset = (page - 1) * per_page

    # Menyusun query SQL dengan filter
    query = 'SELECT * FROM data_excel'
    filters = []
    filter_values = {}

    if location:
        filters.append('(ocr_area LIKE :loc OR ocr_province LIKE :loc OR ocr_city LIKE :loc OR area LIKE :loc)')
        filter_values['loc'] = f'%{location}%'
    
    if filters:
        query += ' WHERE ' + ' AND '.join(filters)
    
    query += ' LIMIT :per_page OFFSET :offset'
    filter_values.update({'per_page': per_page, 'offset': offset})

    # Eksekusi query untuk mendapatkan data
    sql_query = text(query)
    data = pd.read_sql(sql_query, con=db.engine, params=filter_values)

    # Total data dan halaman
    total_query = 'SELECT COUNT(*) as count FROM data_excel'
    if filters:
        total_query += ' WHERE ' + ' AND '.join(filters)
    
    total_sql_query = text(total_query)
    total_data = pd.read_sql(total_sql_query, con=db.engine, params=filter_values)['count'][0]
    total_pages = (total_data // per_page) + (1 if total_data % per_page > 0 else 0)

    # Menentukan range halaman yang ingin ditampilkan (1-10 halaman)
    start_page = max(1, page - 5)
    end_page = min(total_pages, page + 4)

    return render_template(
        'bagi_excel.html',
        data=data.to_dict(orient='records'),  # Mengirimkan data ke template
        page=page,
        total_pages=total_pages,
        start_page=start_page,
        end_page=end_page
    )

@app.route('/bagi-excel-data')
@login_required
def bagi_excel_data():
    page = request.args.get('page', 1, type=int)
    location = request.args.get('location', '')
    search = request.args.get('search', '')  # Ambil search query dari query string
    status_user = request.args.get('status_user', None)  # Ambil status user dari query string
    selected_users = request.args.getlist('selected_users')  # Ambil daftar user yang dipilih

    per_page = 25  # Jumlah item per halaman
    offset = (page - 1) * per_page

    # Menyusun query SQL untuk data Excel dengan filter
    query = 'SELECT * FROM data_excel'
    filters = []
    filter_values = {}

    # Filter lokasi
    if location:
        filters.append('(ocr_area LIKE :loc OR ocr_province LIKE :loc OR ocr_city LIKE :loc OR area LIKE :loc)')
        filter_values['loc'] = f'%{location}%'
    
    # Filter berdasarkan search query (name atau order number)
    if search:
        filters.append('(name LIKE :search OR order_no LIKE :search)')
        filter_values['search'] = f'%{search}%'

    # Filter status user
    if status_user == 'NULL':
        filters.append('nama_user IS NULL')
    
    # Filter berdasarkan user yang dipilih
    if selected_users:
        filters.append('nama_user IN :selected_users')
        filter_values['selected_users'] = tuple(selected_users)  # Menggunakan tuple untuk query SQL

    if filters:
        query += ' WHERE ' + ' AND '.join(filters)
    
    query += ' LIMIT :per_page OFFSET :offset'
    filter_values.update({'per_page': per_page, 'offset': offset})

    # Eksekusi query untuk mendapatkan data
    sql_query = text(query)
    data = pd.read_sql(sql_query, con=db.engine, params=filter_values)

    # Total data dan halaman
    total_query = 'SELECT COUNT(*) as count FROM data_excel'
    if filters:
        total_query += ' WHERE ' + ' AND '.join(filters)
    
    total_sql_query = text(total_query)
    total_data = pd.read_sql(total_sql_query, con=db.engine, params=filter_values)['count'][0]
    total_pages = (total_data // per_page) + (1 if total_data % per_page > 0 else 0)

    # Menentukan range halaman yang ingin ditampilkan (1-10 halaman)
    start_page = max(1, page - 5)
    end_page = min(total_pages, page + 4)

    # Ambil user berdasarkan lokasi
    user_query = 'SELECT * FROM user'
    user_filters = []
    user_filter_values = {}

    if location:
        user_filters.append('lokasi LIKE :loc')
        user_filter_values['loc'] = f'%{location}%'

    if user_filters:
        user_query += ' WHERE ' + ' AND '.join(user_filters)
    
    user_sql_query = text(user_query)
    users = pd.read_sql(user_sql_query, con=db.engine, params=user_filter_values)

    return render_template(
        'bagi_excel.html',
        data=data.to_dict(orient='records'),
        users=users.to_dict(orient='records'),
        page=page,
        total_pages=total_pages,
        start_page=start_page,
        end_page=end_page,
        total_data=len(data)
    )

@app.route('/bagi-excel-data-orang')
@login_required
def bagi_excel_data_orang():
    page = request.args.get('page', 1, type=int)
    location = request.args.get('location', '')  # Ambil lokasi dari query string
    search = request.args.get('search', '')  # Ambil query pencarian
    per_page = 25  # Jumlah item per halaman
    offset = (page - 1) * per_page

    # Menyusun query SQL untuk data Excel dengan filter
    query = 'SELECT * FROM data_excel WHERE nama_user IS NULL'  # Menambahkan filter untuk nama_user
    filters = []
    filter_values = {}

    # Filter lokasi
    if location:
        filters.append('(ocr_area LIKE :loc OR ocr_province LIKE :loc OR ocr_city LIKE :loc OR area LIKE :loc)')
        filter_values['loc'] = f'%{location}%'

    # Menambahkan filter berdasarkan search query (name atau order_no)
    if search:
        filters.append('(name LIKE :search OR order_no LIKE :search)')
        filter_values['search'] = f'%{search}%'

    # Menambahkan filter lokasi ke query jika ada
    if filters:
        query += ' AND ' + ' AND '.join(filters)

    query += ' LIMIT :per_page OFFSET :offset'
    filter_values.update({'per_page': per_page, 'offset': offset})

    # Eksekusi query untuk mendapatkan data
    sql_query = text(query)
    data = pd.read_sql(sql_query, con=db.engine, params=filter_values)

    # Total data dan halaman
    total_query = 'SELECT COUNT(*) as count FROM data_excel WHERE nama_user IS NULL'  # Menghitung total data yang nama_user-nya NULL
    if filters:
        total_query += ' AND ' + ' AND '.join(filters)

    total_sql_query = text(total_query)
    total_data = pd.read_sql(total_sql_query, con=db.engine, params=filter_values)['count'][0]
    total_pages = (total_data // per_page) + (1 if total_data % per_page > 0 else 0)

    # Menentukan range halaman yang ingin ditampilkan (1-10 halaman)
    start_page = max(1, page - 5)
    end_page = min(total_pages, page + 4)

    # Ambil semua user (collector) tanpa filter pada lokasi
    user_query = 'SELECT * FROM user'  # Ambil semua collector tanpa filter
    user_sql_query = text(user_query)
    users = pd.read_sql(user_sql_query, con=db.engine)

    return render_template(
        'bagi_excel_orang.html',
        data=data.to_dict(orient='records'),
        users=users.to_dict(orient='records'),  # Pastikan collectors selalu diambil
        page=page,
        total_pages=total_pages,
        start_page=start_page,
        end_page=end_page,
        total_data=len(data)
    )

@app.route('/apply_users', methods=['POST'])
@login_required
def apply_users():
    selected_users = request.form.getlist('selected_users')
    location = request.form.get('location')

    print(f'Selected Users: {selected_users}, Location: {location}')  # Debugging

    if not selected_users:
        flash('No users selected.')
        return redirect(url_for('view_excel_data'))

    # Ambil data dari data_excel yang sesuai dengan lokasi di semua field yang relevan
    data_to_update = DataExcel.query.filter(
        (DataExcel.ocr_area.like(f'%{location}%')) |
        (DataExcel.ocr_province.like(f'%{location}%')) |
        (DataExcel.ocr_city.like(f'%{location}%')) |
        (DataExcel.area.like(f'%{location}%'))
    ).all()

    total_data = len(data_to_update)
    num_users = len(selected_users)

    print(f'Total data found for location "{location}": {total_data}')  # Log jumlah data

    if total_data == 0:
        flash('No data found for the selected location.')
        return redirect(url_for('view_excel_data'))

    # Bagikan data secara merata hanya untuk data yang sesuai lokasi
    for i, data in enumerate(data_to_update):
        user_index = i % num_users  # Ambil indeks user sesuai dengan data yang sedang diproses
        data.nama_user = selected_users[user_index]  # Assign user ke nama_user
        print(f'Updated {data.id} to nama_user: {data.nama_user} for location {data.ocr_area}')  # Log untuk debug

    # Commit perubahan ke database
    db.session.commit()
    flash(f'Distributed {total_data} records among selected users!')
    return redirect(url_for('view_excel_data'))

@app.route('/apply_users_orang', methods=['POST'])
@login_required
def apply_users_orang():
    selected_users = request.form.getlist('selected_users')
    location = request.form.get('location')

    print(f'Selected Users: {selected_users}, Location: {location}')  # Debugging

    if not selected_users:
        flash('No users selected.')
        return redirect(url_for('bagi_excel_data_orang'))

    # Ambil data dari data_excel yang sesuai dengan lokasi di semua field yang relevan
    data_to_update = DataExcel.query.filter(
        (DataExcel.ocr_area.like(f'%{location}%')) |
        (DataExcel.ocr_province.like(f'%{location}%')) |
        (DataExcel.ocr_city.like(f'%{location}%')) |
        (DataExcel.area.like(f'%{location}%'))
    ).all()

    total_data = len(data_to_update)
    num_users = len(selected_users)

    print(f'Total data found for location "{location}": {total_data}')  # Log jumlah data

    if total_data == 0:
        flash('No data found for the selected location.')
        return redirect(url_for('bagi_excel_data_orang'))

    # Bagikan data secara merata hanya untuk data yang sesuai lokasi
    for i, data in enumerate(data_to_update):
        user_index = i % num_users  # Ambil indeks user sesuai dengan data yang sedang diproses
        data.nama_user = selected_users[user_index]  # Assign user ke nama_user
        print(f'Updated {data.id} to nama_user: {data.nama_user} for location {data.ocr_area}')  # Log untuk debug

    # Commit perubahan ke database
    db.session.commit()
    flash(f'Distributed {total_data} records among selected users!')
    return redirect(url_for('bagi_excel_data_orang'))

@app.route('/apply-collector', methods=['POST'])
def apply_collector():
    selected_orders = request.form.get('selected_orders').split(',')
    selected_user = request.form.get('selected_user')

    users = User.query.filter(User.role != 'admin').all()

    # Update field 'nama_user' di tabel 'data_excel'
    for order_no in selected_orders:
        data = DataExcel.query.filter_by(order_no=order_no).first()
        if data:
            data.nama_user = selected_user
            db.session.commit()

    flash(f'Collector {selected_user} has been successfully applied to selected orders!', 'success')
    return redirect(url_for('bagi_excel_data_orang'))

if __name__ == '__main__':
    # Create tables if not exists
    app.run(debug=True) 
    with app.app_context():
        db.create_all()
    app.run(host='0.0.0.0', port=5000)
