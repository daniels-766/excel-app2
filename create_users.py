from app import app, db, User, bcrypt

# Buat konteks aplikasi
with app.app_context():
    # Create admin user
    admin = User(username='Willy J Leasiwal', password_hash=bcrypt.generate_password_hash('willy@888').decode('utf-8'), role='admin')
    db.session.add(admin)

    # Commit changes to the database
    db.session.commit()

    print("Users created successfully!")
