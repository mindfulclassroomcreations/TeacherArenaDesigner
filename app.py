from flask import Flask, render_template, request, send_file, jsonify, Response, stream_with_context, redirect, url_for, flash
import os
import shutil
import tempfile
import zipfile
import json
from werkzeug.utils import secure_filename
from worksheet_generator import generate_worksheets
from caterpillar_generator import generate_caterpillar_worksheets
import traceback
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from flask_bcrypt import Bcrypt
from models import db, User, Config

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'your-secret-key-change-this')
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///database.db').replace('postgres://', 'postgresql://', 1)

db.init_app(app)
bcrypt = Bcrypt(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def zip_directory(folder_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

# --- Auth Routes ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        if user and bcrypt.check_password_hash(user.password_hash, password):
            login_user(user)
            return redirect(url_for('admin_dashboard'))
        else:
            flash('Login Unsuccessful. Please check username and password')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('landing'))

@app.route('/admin', methods=['GET'])
@login_required
def admin_dashboard():
    if not current_user.is_admin:
        return redirect(url_for('landing'))
    
    api_key_config = Config.query.filter_by(key_name='openai_api_key').first()
    api_key = api_key_config.value if api_key_config else ''
    users = User.query.all()
    return render_template('admin.html', api_key=api_key, users=users)

@app.route('/admin/update_key', methods=['POST'])
@login_required
def update_key():
    if not current_user.is_admin:
        return redirect(url_for('landing'))
    
    new_key = request.form.get('api_key')
    config = Config.query.filter_by(key_name='openai_api_key').first()
    if config:
        config.value = new_key
    else:
        config = Config(key_name='openai_api_key', value=new_key)
        db.session.add(config)
    db.session.commit()
    flash('API Key Updated!')
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/add_user', methods=['POST'])
@login_required
def add_user():
    if not current_user.is_admin:
        return redirect(url_for('landing'))
    
    username = request.form.get('new_username')
    password = request.form.get('new_password')
    hashed_password = bcrypt.generate_password_hash(password).decode('utf-8')
    new_user = User(username=username, password_hash=hashed_password)
    try:
        db.session.add(new_user)
        db.session.commit()
        flash('User Added!')
    except:
        flash('Username already exists.')
    return redirect(url_for('admin_dashboard'))

# --- Main Routes ---
@app.route('/')
def landing():
    return render_template('landing.html')

@app.route('/academy-ready')
def academy_ready():
    return render_template('academy_ready.html')

@app.route('/dreaming-caterpillar')
def dreaming_caterpillar():
    return render_template('caterpillar.html')

def handle_generation(generator_func, request):
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    # Fetch API Key from DB
    api_key_config = Config.query.filter_by(key_name='openai_api_key').first()
    if not api_key_config or not api_key_config.value:
        return jsonify({'error': 'OpenAI API Key not set by Admin.'}), 400
    
    api_key = api_key_config.value
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsx'):
        # Create a temporary directory for this generation session
        session_dir = tempfile.mkdtemp()
        upload_path = os.path.join(session_dir, secure_filename(file.filename))
        file.save(upload_path)
        
        output_dir = os.path.join(session_dir, 'output')
        os.makedirs(output_dir, exist_ok=True)

        def generate_stream():
            try:
                # Run generation
                for update in generator_func(upload_path, output_dir, api_key):
                    if update['type'] == 'progress':
                        yield f"data: {json.dumps(update)}\n\n"
                    
                    elif update['type'] == 'result':
                        # Zip the individual topic folder
                        topic_path = update['path']
                        topic_name = update['topic']
                        safe_topic_name = secure_filename(topic_name)
                        zip_filename = f"{safe_topic_name}.zip"
                        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                        zip_directory(topic_path, zip_path)
                        
                        update['download_url'] = f'/download/{zip_filename}'
                        yield f"data: {json.dumps(update)}\n\n"
                    
                    elif update['type'] == 'complete':
                        # Zip the full output
                        full_output_path = update['path']
                        zip_filename = f"worksheets_{os.path.basename(session_dir)}.zip"
                        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
                        zip_directory(full_output_path, zip_path)
                        
                        update['download_url'] = f'/download/{zip_filename}'
                        yield f"data: {json.dumps(update)}\n\n"

                # Cleanup session dir
                shutil.rmtree(session_dir)

            except Exception as e:
                traceback.print_exc()
                error_msg = {'type': 'error', 'message': str(e)}
                yield f"data: {json.dumps(error_msg)}\n\n"

        return Response(stream_with_context(generate_stream()), mimetype='text/event-stream')

    else:
        return jsonify({'error': 'Invalid file type. Please upload an Excel (.xlsx) file.'}), 400

@app.route('/generate-academy', methods=['POST'])
def generate_academy():
    return handle_generation(generate_worksheets, request)

@app.route('/generate-caterpillar', methods=['POST'])
def generate_caterpillar():
    return handle_generation(generate_caterpillar_worksheets, request)

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

# Initialize database tables
with app.app_context():
    db.create_all()
    # Create default admin if not exists
    if not User.query.filter_by(username='admin').first():
        hashed_pw = bcrypt.generate_password_hash('admin123').decode('utf-8')
        admin = User(username='admin', password_hash=hashed_pw, is_admin=True)
        db.session.add(admin)
        db.session.commit()
        print("Default admin created: admin / admin123")

if __name__ == '__main__':
    app.run(debug=True, port=5000)
