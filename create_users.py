from app import app, db, User, bcrypt

# Buat konteks aplikasi
with app.app_context():
    # Create admin user
    admin = User(username='admin', password_hash=bcrypt.generate_password_hash('12345678').decode('utf-8'), role='admin')
    db.session.add(admin)

    # Commit changes to the database
    db.session.commit()

    print("Users created successfully!")
