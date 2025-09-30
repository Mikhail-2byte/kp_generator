# app/database.py
import sqlite3
import logging
from datetime import datetime
from contextlib import closing

def connect_db():
    """Устанавливает соединение с базой данных"""
    return sqlite3.connect('kp_generator.db', detect_types=sqlite3.PARSE_DECLTYPES)

def init_db():
    """Инициализирует базу данных"""
    schema_sql_content = """
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        last_login DATETIME
    );

    CREATE TABLE IF NOT EXISTS generation_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
        tender_number TEXT,
        company TEXT NOT NULL,
        product TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        cost_price REAL NOT NULL,
        weight REAL NOT NULL,
        logistics REAL NOT NULL,
        margin_percent REAL NOT NULL,
        final_price REAL NOT NULL,
        drawing_number TEXT,
        material TEXT,
        delivery_address TEXT,
        duty_percent REAL DEFAULT 0,
        delivery_time INTEGER DEFAULT 0,
        comment TEXT,
        user_id INTEGER,
        FOREIGN KEY(user_id) REFERENCES users(id)
    );

    CREATE INDEX IF NOT EXISTS idx_generation_history_timestamp ON generation_history(timestamp);
    CREATE INDEX IF NOT EXISTS idx_generation_history_company ON generation_history(company);
    """

    with closing(connect_db()) as db:
        cursor = db.cursor()
        cursor.executescript(schema_sql_content)

        cursor.execute("PRAGMA table_info(generation_history)")
        columns = [row[1] for row in cursor.fetchall()]
        if 'comment' not in columns:
            cursor.execute('ALTER TABLE generation_history ADD COLUMN comment TEXT')
        if 'user_id' not in columns:
            cursor.execute('ALTER TABLE generation_history ADD COLUMN user_id INTEGER')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_generation_history_user ON generation_history(user_id)')

        cursor.execute("PRAGMA table_info(users)")
        user_columns = [row[1] for row in cursor.fetchall()]
        if 'last_login' not in user_columns:
            cursor.execute('ALTER TABLE users ADD COLUMN last_login DATETIME')

        db.commit()

def get_generation_history(config):
    """Получает историю генераций из базы данных"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            
            cursor.execute('''
                SELECT gh.id, gh.timestamp, gh.tender_number, gh.company, gh.product,
                       gh.margin_percent, gh.final_price, gh.drawing_number, u.username
                FROM generation_history gh
                LEFT JOIN users u ON gh.user_id = u.id
                ORDER BY gh.timestamp DESC
                LIMIT ?
            ''', (config.get('max_history_items', 50),))
            
            history = cursor.fetchall()
            
            result = []
            for item in history:
                timestamp_value = item[1]
                if isinstance(timestamp_value, str):
                    try:
                        formatted_ts = datetime.strptime(timestamp_value, '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M')
                    except ValueError:
                        formatted_ts = timestamp_value
                else:
                    formatted_ts = timestamp_value.strftime('%d.%m.%Y %H:%M')

                result.append({
                    'id': item[0],
                    'timestamp': formatted_ts,
                    'tender_number': item[2] or 'Не указан',
                    'company': item[3],
                    'product': item[4],
                    'margin_percent': item[5],
                    'final_price': item[6],
                    'drawing_number': item[7] or 'Не указан',
                    'username': item[8]
                })
            
            return result
    except Exception as e:
        logging.error(f'Error getting generation history: {str(e)}')
        return []

def save_generation_history(form_data, final_price, config, user_id=None):
    """Сохраняет данные о генерации в базу данных"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            
            cursor.execute('''
                INSERT INTO generation_history 
                (tender_number, company, product, quantity, cost_price, weight, logistics, margin_percent, final_price,
                 drawing_number, material, delivery_address, duty_percent, delivery_time, comment, user_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                form_data.get('tender_number', ''),
                form_data.get('company', ''),
                form_data.get('product', ''),
                int(form_data.get('quantity', 0)),
                float(form_data.get('cost_price', 0)),
                float(form_data.get('weight', 0)),
                float(form_data.get('logistics', 0)),
                float(form_data.get('margin_percent', config.get('margin_percent', 30))),
                final_price,
                form_data.get('drawing_number', ''),
                form_data.get('material', ''),
                form_data.get('delivery_address', ''),
                float(form_data.get('duty_percent', config.get('default_duty_percent', 0))),
                int(form_data.get('delivery_time', 0)),
                form_data.get('comment', ''),
                user_id
            ))
            
            db.commit()
        return True
    except Exception as e:
        logging.error(f'Error saving generation history: {str(e)}')
        return False

def get_generation_details(record_id):
    """Получает детальную информацию о конкретной записи"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('''
                SELECT gh.*, u.username
                FROM generation_history gh
                LEFT JOIN users u ON gh.user_id = u.id
                WHERE gh.id = ?
            ''', (record_id,))
            record = cursor.fetchone()
            
            if record:
                columns = ['id', 'timestamp', 'tender_number', 'company', 'product', 'quantity', 
                          'cost_price', 'weight', 'logistics', 'margin_percent', 'final_price',
                          'drawing_number', 'material', 'delivery_address', 'duty_percent', 'delivery_time', 'comment', 'user_id', 'username']
                return dict(zip(columns, record))
            return None
    except Exception as e:
        logging.error(f'Error getting generation details: {str(e)}')
        return None

def load_generation_data(gen_id):
    """Загружает данные конкретной генерации для повторного использования"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('SELECT * FROM generation_history WHERE id = ?', (gen_id,))
            generation = cursor.fetchone()
        
        if generation:
            columns = ['id', 'timestamp', 'tender_number', 'company', 'product', 'quantity', 
                      'cost_price', 'weight', 'logistics', 'margin_percent', 'final_price',
                      'drawing_number', 'material', 'delivery_address', 'duty_percent', 'delivery_time', 'comment', 'user_id']
            return dict(zip(columns, generation))
        return None
    except Exception as e:
        logging.error(f'Error loading generation data: {str(e)}')
        return None


def create_user(username, password_hash):
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('INSERT INTO users (username, password_hash) VALUES (?, ?)', (username, password_hash))
            db.commit()
            return cursor.lastrowid
    except sqlite3.IntegrityError:
        logging.error(f'User with username {username} already exists')
        return None
    except Exception as e:
        logging.error(f'Error creating user: {str(e)}')
        return None


def get_user_by_username(username):
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('SELECT id, username, password_hash, created_at, last_login FROM users WHERE username = ?', (username,))
            return cursor.fetchone()
    except Exception as e:
        logging.error(f'Error fetching user by username: {str(e)}')
        return None


def get_user_by_id(user_id):
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('SELECT id, username, password_hash, created_at, last_login FROM users WHERE id = ?', (user_id,))
            return cursor.fetchone()
    except Exception as e:
        logging.error(f'Error fetching user by id: {str(e)}')
        return None


def update_last_login(user_id, last_login=None):
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('UPDATE users SET last_login = COALESCE(?, CURRENT_TIMESTAMP) WHERE id = ?', (last_login, user_id))
            db.commit()
            return cursor.rowcount > 0
    except Exception as e:
        logging.error(f'Error updating last login: {str(e)}')
        return False


def get_user_statistics(user_id):
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            cursor.execute('''
                SELECT COUNT(*), COALESCE(MAX(timestamp), '')
                FROM generation_history
                WHERE user_id = ?
            ''', (user_id,))
            count, last_timestamp = cursor.fetchone()

            cursor.execute('''
                SELECT id, company, product, margin_percent, final_price, timestamp
                FROM generation_history
                WHERE user_id = ?
                ORDER BY timestamp DESC
                LIMIT 5
            ''', (user_id,))
            recent = cursor.fetchall()

            return {
                'total_generations': count or 0,
                'last_generation_at': last_timestamp,
                'recent_generations': recent
            }
    except Exception as e:
        logging.error(f'Error getting user statistics: {str(e)}')
        return {
            'total_generations': 0,
            'last_generation_at': None,
            'recent_generations': []
        }