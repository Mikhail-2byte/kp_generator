
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

CREATE INDEX IF NOT EXISTS idx_generation_history_timestamp ON generation_history(timestamp);
CREATE INDEX IF NOT EXISTS idx_generation_history_company ON generation_history(company);
