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
        delivery_time INTEGER DEFAULT 0
    );
    """
    
    with closing(connect_db()) as db:
        db.cursor().executescript(schema_sql_content)
        db.commit()

def get_generation_history(config):
    """Получает историю генераций из базы данных"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            
            cursor.execute('''
                SELECT id, timestamp, tender_number, company, product, margin_percent, final_price
                FROM generation_history 
                ORDER BY timestamp DESC
                LIMIT ?
            ''', (config.get('max_history_items', 50),))
            
            history = cursor.fetchall()
            
            result = []
            for item in history:
                result.append({
                    'id': item[0],
                    'timestamp': datetime.strptime(item[1], '%Y-%m-%d %H:%M:%S').strftime('%d.%m.%Y %H:%M'),
                    'tender_number': item[2] or 'Не указан',
                    'company': item[3],
                    'product': item[4],
                    'margin_percent': item[5],
                    'final_price': item[6]
                })
            
            return result
    except Exception as e:
        logging.error(f'Error getting generation history: {str(e)}')
        return []

def save_generation_history(form_data, final_price, config):
    """Сохраняет данные о генерации в базу данных"""
    try:
        with closing(connect_db()) as db:
            cursor = db.cursor()
            
            cursor.execute('''
                INSERT INTO generation_history 
                (tender_number, company, product, quantity, cost_price, weight, logistics, margin_percent, final_price,
                 drawing_number, material, delivery_address, duty_percent, delivery_time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                int(form_data.get('delivery_time', 0))
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
            cursor.execute('SELECT * FROM generation_history WHERE id = ?', (record_id,))
            record = cursor.fetchone()
            
            if record:
                columns = ['id', 'timestamp', 'tender_number', 'company', 'product', 'quantity', 
                          'cost_price', 'weight', 'logistics', 'margin_percent', 'final_price',
                          'drawing_number', 'material', 'delivery_address', 'duty_percent', 'delivery_time']
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
                      'drawing_number', 'material', 'delivery_address', 'duty_percent', 'delivery_time']
            return dict(zip(columns, generation))
        return None
    except Exception as e:
        logging.error(f'Error loading generation data: {str(e)}')
        return None