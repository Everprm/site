import sqlite3
import psycopg2
from psycopg2 import sql

# --- НАСТРОЙКИ ---
SQLITE_DB_PATH = 'warehouse.db'

# Строка подключения к Timeweb Cloud PostgreSQL
# БЕЗ слова psql, БЕЗ кавычек вокруг — просто строка Python!
DATABASE_URL = 'postgresql://gen_user:uSa7%3F(EecL%7B2%3B4@201.34.136.110:5432/default_db?sslmode=require'


def migrate_data():
    print("🚀 Начинаем миграцию данных...")

    # 1. Подключаемся к SQLite
    try:
        sqlite_conn = sqlite3.connect(SQLITE_DB_PATH)
        sqlite_conn.row_factory = sqlite3.Row
        sqlite_cur = sqlite_conn.cursor()
        print(f"✅ Подключение к SQLite ({SQLITE_DB_PATH}) установлено.")
    except Exception as e:
        print(f"❌ Ошибка подключения к SQLite: {e}")
        return

    # 2. Подключаемся к PostgreSQL
    try:
        pg_conn = psycopg2.connect(DATABASE_URL)
        pg_cur = pg_conn.cursor()
        print("✅ Подключение к PostgreSQL установлено.")
    except Exception as e:
        print(f"❌ Ошибка подключения к PostgreSQL: {e}")
        sqlite_conn.close()
        return

    tables_to_migrate = ['users', 'products', 'stock_moves', 'reserves']

    try:
        for table in tables_to_migrate:
            print(f"\n📦 Перенос таблицы: {table}")

            sqlite_cur.execute(f"SELECT * FROM {table}")
            rows = sqlite_cur.fetchall()
            if not rows:
                print(f"   ⚠️ Таблица {table} пуста, пропускаем.")
                continue

            column_names = [description[0] for description in sqlite_cur.description]
            print(f"   Найдено строк: {len(rows)}")

            # Очищаем таблицу в PostgreSQL
            pg_cur.execute(sql.SQL("TRUNCATE TABLE {} RESTART IDENTITY CASCADE").format(sql.Identifier(table)))
            pg_conn.commit()
            print(f"   🗑️ Таблица {table} в PostgreSQL очищена.")

            # Вставляем данные
            placeholders = ', '.join(['%s'] * len(column_names))
            columns = ', '.join([f'"{col}"' for col in column_names])
            insert_query = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
                sql.Identifier(table),
                sql.SQL(columns),
                sql.SQL(placeholders)
            )

            data_to_insert = [tuple(row) for row in rows]
            pg_cur.executemany(insert_query, data_to_insert)
            pg_conn.commit()
            print(f"   ✅ Перенесено строк: {len(data_to_insert)}")

        # Обновляем последовательности ID
        print("\n🔢 Обновляем последовательности ID...")
        for table in tables_to_migrate:
            try:
                seq_name = f"{table}_id_seq"
                pg_cur.execute(f"SELECT setval('{seq_name}', COALESCE((SELECT MAX(id) FROM {table}), 1))")
                pg_conn.commit()
                print(f"   ✅ Последовательность для {table} обновлена.")
            except Exception as e:
                pg_conn.rollback()
                print(f"   ⚠️ Не удалось обновить последовательность для {table}: {e}")

        print("\n🎉 Миграция данных успешно завершена!")

    except Exception as e:
        print(f"\n❌ Ошибка во время миграции: {e}")
        pg_conn.rollback()
    finally:
        sqlite_conn.close()
        pg_conn.close()
        print("🔒 Соединения закрыты.")


if __name__ == '__main__':
    migrate_data()