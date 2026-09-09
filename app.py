from flask import Flask, request, jsonify, render_template, session, redirect, url_for, send_file
import sqlite3
import bcrypt
from datetime import datetime, timedelta
import os
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from functools import wraps

app = Flask(__name__)
# Секретный ключ для сессий
app.secret_key = 'supersecretkey123!@#$%'
app.permanent_session_lifetime = timedelta(hours=8)


# ============================================================
# ФУНКЦИИ РАБОТЫ С БАЗОЙ ДАННЫХ
# ============================================================

def get_db():
    """Подключение к базе данных"""
    conn = sqlite3.connect('warehouse.db')
    conn.row_factory = sqlite3.Row
    return conn


# ============================================================
# ДЕКОРАТОРЫ ДЛЯ ПРОВЕРКИ ПРАВ ДОСТУПА
# ============================================================

def login_required(f):
    """Декоратор: только для авторизованных пользователей"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated_function


def role_required(allowed_roles):
    """Декоратор: только для пользователей с определенной ролью"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'role' not in session or session['role'] not in allowed_roles:
                return jsonify({'error': 'Доступ запрещен. Недостаточно прав.'}), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator


# ============================================================
# СТРАНИЦЫ
# ============================================================

@app.route('/login')
def login_page():
    """Страница входа"""
    if 'user_id' in session:
        return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/')
@login_required
def index():
    """Главная страница"""
    return render_template('index.html', username=session.get('username'), role=session.get('role'))


@app.route('/history')
@login_required
def history():
    """Страница журнала действий (только для админа)"""
    if session.get('role') != 'admin':
        return render_template('access_denied.html'), 403
    
    conn = get_db()
    logs = conn.execute('''
        SELECT 
            sm.id,
            p.name as product_name,
            sm.quantity,
            sm.user,
            sm.comment,
            sm.created_at
        FROM stock_moves sm
        JOIN products p ON sm.product_id = p.id
        ORDER BY sm.created_at DESC
        LIMIT 200
    ''').fetchall()
    conn.close()
    return render_template('history.html', logs=logs, username=session.get('username'))


@app.route('/logout')
def logout():
    """Выход из системы"""
    session.clear()
    return redirect(url_for('login_page'))


# ============================================================
# API - АВТОРИЗАЦИЯ
# ============================================================

@app.route('/api/login', methods=['POST'])
def api_login():
    """API: Вход пользователя"""
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    
    if not username or not password:
        return jsonify({'error': 'Введите логин и пароль'}), 400
    
    conn = get_db()
    user = conn.execute(
        "SELECT id, name, password_hash, role FROM users WHERE name = ?",
        (username,)
    ).fetchone()
    conn.close()
    
    if not user:
        return jsonify({'error': 'Пользователь не найден'}), 401
    
    # Проверяем пароль
    try:
        if bcrypt.checkpw(password.encode('utf-8'), user['password_hash'].encode('utf-8')):
            session.permanent = True
            session['user_id'] = user['id']
            session['username'] = user['name']
            session['role'] = user['role']
            return jsonify({'status': 'OK', 'user': user['name'], 'role': user['role']})
        else:
            return jsonify({'error': 'Неверный пароль'}), 401
    except Exception as e:
        return jsonify({'error': f'Ошибка проверки пароля: {str(e)}'}), 500


@app.route('/api/current_user')
@login_required
def current_user():
    """API: Получить информацию о текущем пользователе"""
    return jsonify({
        'username': session.get('username'),
        'role': session.get('role')
    })


@app.route('/api/users')
@login_required
def get_users():
    """API: Список всех пользователей"""
    conn = get_db()
    users = conn.execute("SELECT name FROM users ORDER BY name").fetchall()
    conn.close()
    return jsonify([u['name'] for u in users])


# ============================================================
# API - ОСТАТКИ И ДВИЖЕНИЯ
# ============================================================

@app.route('/api/balances')
@login_required
def balances():
    """API: Получить текущие остатки всех товаров"""
    conn = get_db()
    data = conn.execute('''
        SELECT 
            p.id, 
            p.name, 
            p.unit, 
            COALESCE(SUM(sm.quantity), 0) as balance
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id
        ORDER BY p.id
    ''').fetchall()
    conn.close()
    return jsonify([dict(row) for row in data])


@app.route('/api/ship', methods=['POST'])
@login_required
def ship():
    """API: Отгрузка товара (уменьшает остаток)"""
    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    user = session.get('username', 'Неизвестный')
    comment = data.get('comment', 'Отгрузка')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    
    # Проверяем остаток
    check = conn.execute('''
        SELECT COALESCE(SUM(quantity), 0) as total 
        FROM stock_moves 
        WHERE product_id = ?
    ''', (prod_id,)).fetchone()
    
    current_balance = check['total']
    if current_balance < qty:
        conn.close()
        return jsonify({'error': f'Недостаточно товара! Остаток: {current_balance}'}), 400

    # Записываем отгрузку со знаком МИНУС
    conn.execute(
        "INSERT INTO stock_moves (product_id, quantity, user, comment) VALUES (?, ?, ?, ?)",
        (prod_id, -abs(qty), user, comment)
    )
    conn.commit()
    conn.close()
    return jsonify({'status': 'OK', 'message': f'Отгружено {qty} единиц'})


@app.route('/api/receive', methods=['POST'])
@login_required
@role_required(['manager', 'admin'])
def receive():
    """API: Пополнение склада (увеличивает остаток). Только для manager и admin"""
    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    user = session.get('username', 'Неизвестный')
    comment = data.get('comment', 'Пополнение')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    
    # Записываем пополнение со знаком ПЛЮС
    conn.execute(
        "INSERT INTO stock_moves (product_id, quantity, user, comment) VALUES (?, ?, ?, ?)",
        (prod_id, abs(qty), user, comment)
    )
    conn.commit()
    conn.close()
    return jsonify({'status': 'OK', 'message': f'Добавлено {qty} единиц'})


# ============================================================
# ВЫГРУЗКА В EXCEL (ТОЛЬКО ДЛЯ АДМИНА)
# ============================================================

@app.route('/export_excel')
@login_required
def export_excel():
    """Экспорт актуальных остатков в Excel (только для админа)"""
    # Проверяем, что пользователь админ
    if session.get('role') != 'admin':
        return render_template('access_denied.html'), 403
    
    conn = get_db()
    
    # Получаем все товары с их текущими остатками
    data = conn.execute('''
        SELECT 
            p.id,
            p.name,
            COALESCE(SUM(sm.quantity), 0) as balance,
            p.unit
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id
        ORDER BY p.id
    ''').fetchall()
    conn.close()
    
    # --- СОЗДАЕМ EXCEL-ФАЙЛ ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Остатки склада"
    
    # --- СТИЛИ ---
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1a3c5e", end_color="1a3c5e", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    
    data_font = Font(size=10)
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    number_alignment = Alignment(horizontal="center", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # --- ЗАГОЛОВКИ ---
    headers = ['№ п/п', 'Наименование позиции', 'Ед. изм.', 'Текущий остаток, шт']
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # --- ЗАПОЛНЯЕМ ДАННЫЕ ---
    for row_idx, item in enumerate(data, 2):
        # № п/п
        cell = ws.cell(row=row_idx, column=1, value=item['id'])
        cell.font = data_font
        cell.alignment = number_alignment
        cell.border = thin_border
        
        # Наименование
        cell = ws.cell(row=row_idx, column=2, value=item['name'])
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = thin_border
        
        # Ед. изм.
        cell = ws.cell(row=row_idx, column=3, value=item['unit'] or 'шт')
        cell.font = data_font
        cell.alignment = number_alignment
        cell.border = thin_border
        
        # Остаток
        cell = ws.cell(row=row_idx, column=4, value=item['balance'])
        cell.alignment = number_alignment
        cell.border = thin_border
        
        # Выделяем цветом в зависимости от остатка
        if item['balance'] < 5 and item['balance'] > 0:
            # Мало остатка — желтый
            cell.font = Font(color="FF8F00", bold=True, size=10)
            cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        elif item['balance'] <= 0:
            # Нет в наличии — красный
            cell.font = Font(color="FF0000", bold=True, size=10)
            cell.fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
            cell.value = f"{item['balance']} (НЕТ В НАЛИЧИИ!)"
        else:
            # Нормальный остаток
            cell.font = Font(color="1B5E20", size=10)
    
    # --- АВТОПОДБОР ШИРИНЫ КОЛОНОК ---
    column_widths = {
        'A': 10,   # №
        'B': 70,   # Наименование
        'C': 14,   # Ед. изм.
        'D': 25    # Остаток
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # --- ЗАМОРОЗКА ПЕРВОЙ СТРОКИ ---
    ws.freeze_panes = 'A2'
    
    # --- СТАТИСТИКА ВНИЗУ ---
    total_items = len(data)
    low_items = len([item for item in data if 0 < item['balance'] < 5])
    zero_items = len([item for item in data if item['balance'] <= 0])
    
    # Пустая строка
    ws.append([])
    
    # Информация о выгрузке
    ws.append([f'Дата выгрузки: {datetime.now().strftime("%d.%m.%Y %H:%M")}'])
    ws.append([f'Всего позиций: {total_items}'])
    ws.append([f'Позиций с остатком менее 5 шт: {low_items}'])
    ws.append([f'Позиций с нулевым остатком: {zero_items}'])
    ws.append([f'Выгрузил: {session.get("username")} (администратор)'])
    
    # --- СОХРАНЯЕМ В БУФЕР ---
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    # --- ОТПРАВЛЯЕМ ФАЙЛ ---
    filename = f"Остатки_склада_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ============================================================
# ЗАПУСК ПРИЛОЖЕНИЯ
# ============================================================

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)