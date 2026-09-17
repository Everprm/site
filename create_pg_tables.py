import psycopg2

DATABASE_URL = 'postgresql://gen_user:uSa7%3F(EecL%7B2%3B4@201.34.136.110:5432/default_db?sslmode=require'

# SQL для создания таблиц в PostgreSQL
SQL = """
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'worker'
);

CREATE TABLE IF NOT EXISTS products (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL,
    unit TEXT DEFAULT 'шт'
);

CREATE TABLE IF NOT EXISTS stock_moves (
    id SERIAL PRIMARY KEY,
    product_id INTEGER,
    quantity INTEGER,
    "user" TEXT DEFAULT 'Неизвестный',
    comment TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS reserves (
    id SERIAL PRIMARY KEY,
    product_id INTEGER NOT NULL,
    quantity INTEGER NOT NULL,
    "user" TEXT NOT NULL,
    comment TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    is_active INTEGER DEFAULT 1
);
"""

print("🚀 Создаём таблицы в PostgreSQL...")

try:
    conn = psycopg2.connect(DATABASE_URL)
    cur = conn.cursor()
    cur.execute(SQL)
    conn.commit()
    
    # Проверяем, что таблицы созданы
    cur.execute("""
        SELECT table_name 
        FROM information_schema.tables 
        WHERE table_schema = 'public' 
        ORDER BY table_name
    """)
    tables = [row[0] for row in cur.fetchall()]
    
    print("✅ Таблицы успешно созданы!")
    print(f"📋 Список таблиц в БД: {tables}")
    
    cur.close()
    conn.close()
    
except Exception as e:
    print(f"❌ Ошибка: {e}")