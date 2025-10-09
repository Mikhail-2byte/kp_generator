
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    last_login DATETIME,
    last_name TEXT,
    first_name TEXT,
    contact_info TEXT,
    role TEXT NOT NULL DEFAULT 'user'
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
CREATE INDEX IF NOT EXISTS idx_generation_history_user ON generation_history(user_id);
