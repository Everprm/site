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
app.secret_key = 'supersecretkey123!@#$%'
app.permanent_session_lifetime = timedelta(hours=8)


# ============================================================
# ПЕРМСКОЕ ВРЕМЯ (UTC+5)
# ============================================================

def get_perm_time():
    """Возвращает текущее пермское время (UTC+5)"""
    return datetime.utcnow() + timedelta(hours=5)


# ============================================================
# БАЗА ДАННЫХ
# ============================================================

def get_db():
    conn = sqlite3.connect('warehouse.db')
    conn.row_factory = sqlite3.Row
    return conn


# ============================================================
# ДЕКОРАТОРЫ
# ============================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login_page'))
        return f(*args, **kwargs)
    return decorated_function


def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'role' not in session or session['role'] not in allowed_roles:
                return jsonify({'error': 'Доступ запрещен. Недостаточно прав.'}), 403
            return f(*args, **kwargs)
        return decorated_function
    return decorator


# ============================================================
# ПРАВА ДОСТУПА ПО ИМЕНАМ
# ============================================================

# Кто МОЖЕТ отгружать
CAN_SHIP_USERS = ['Павел', 'Валерий']

# Кто МОЖЕТ резервировать и снимать резерв
CAN_RESERVE_USERS = ['Павел', 'Евгений', 'Виталий']


def can_ship(username):
    return username in CAN_SHIP_USERS


def can_reserve(username):
    return username in CAN_RESERVE_USERS


# ============================================================
# СТРАНИЦЫ
# ============================================================

@app.route('/login')
def login_page():
    if 'user_id' in session:
        return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/')
@login_required
def index():
    return render_template('index.html', username=session.get('username'), role=session.get('role'))


@app.route('/history')
@login_required
def history():
    """Журнал (только для админа). Время — пермское."""
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
            datetime(sm.created_at, '+5 hours') as created_at
        FROM stock_moves sm
        JOIN products p ON sm.product_id = p.id
        ORDER BY sm.created_at DESC
        LIMIT 200
    ''').fetchall()
    conn.close()
    return render_template('history.html', logs=logs, username=session.get('username'))


@app.route('/reserves')
@login_required
def reserves_page():
    return render_template('reserves.html', username=session.get('username'))


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))


# ============================================================
# API - АВТОРИЗАЦИЯ
# ============================================================

@app.route('/api/login', methods=['POST'])
def api_login():
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
    username = session.get('username')
    return jsonify({
        'username': username,
        'role': session.get('role'),
        'can_ship': can_ship(username),
        'can_reserve': can_reserve(username)
    })


@app.route('/api/users')
@login_required
def get_users():
    conn = get_db()
    users = conn.execute("SELECT name FROM users ORDER BY name").fetchall()
    conn.close()
    return jsonify([u['name'] for u in users])


# ============================================================
# API - ОСТАТКИ
# ============================================================

@app.route('/api/balances')
@login_required
def balances():
    conn = get_db()
    data = conn.execute('''
        SELECT 
            p.id, 
            p.name, 
            p.unit, 
            COALESCE(SUM(sm.quantity), 0) as balance,
            COALESCE((
                SELECT SUM(r.quantity) 
                FROM reserves r 
                WHERE r.product_id = p.id AND r.is_active = 1
            ), 0) as reserved
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id
        ORDER BY p.id
    ''').fetchall()
    conn.close()
    
    result = []
    for row in data:
        item = dict(row)
        item['available'] = item['balance'] - item['reserved']
        result.append(item)
    
    return jsonify(result)


@app.route('/api/ship', methods=['POST'])
@login_required
def ship():
    """Отгрузка. Только для Павла и Валерия."""
    username = session.get('username')
    
    if not can_ship(username):
        return jsonify({'error': f'❌ У пользователя {username} нет прав на отгрузку'}), 403

    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    user = username
    comment = data.get('comment', 'Отгрузка')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    
    check = conn.execute('''
        SELECT 
            COALESCE((SELECT SUM(quantity) FROM stock_moves WHERE product_id = ?), 0) as balance,
            COALESCE((SELECT SUM(quantity) FROM reserves WHERE product_id = ? AND is_active = 1), 0) as reserved
    ''', (prod_id, prod_id)).fetchone()
    
    balance = check['balance']
    reserved = check['reserved']
    available = balance - reserved
    
    if qty > available:
        conn.close()
        if reserved > 0:
            return jsonify({
                'error': f'❌ Нельзя отгрузить {qty} шт! Доступно: {available} (остаток: {balance}, в резерве: {reserved})'
            }), 400
        else:
            return jsonify({
                'error': f'❌ Недостаточно товара! Остаток: {balance}'
            }), 400

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
    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    user = session.get('username', 'Неизвестный')
    comment = data.get('comment', 'Пополнение')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    conn.execute(
        "INSERT INTO stock_moves (product_id, quantity, user, comment) VALUES (?, ?, ?, ?)",
        (prod_id, abs(qty), user, comment)
    )
    conn.commit()
    conn.close()
    return jsonify({'status': 'OK', 'message': f'Добавлено {qty} единиц'})


# ============================================================
# API - РЕЗЕРВЫ (Павел, Евгений, Виталий)
# ============================================================

@app.route('/api/reserve', methods=['POST'])
@login_required
def reserve_product():
    username = session.get('username')
    
    if not can_reserve(username):
        return jsonify({'error': f'❌ У пользователя {username} нет прав на резервирование'}), 403

    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    user = username
    comment = (data.get('comment') or 'Без пометки').strip() or 'Без пометки'

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    
    check = conn.execute('''
        SELECT 
            COALESCE((SELECT SUM(quantity) FROM stock_moves WHERE product_id = ?), 0) as balance,
            COALESCE((SELECT SUM(quantity) FROM reserves WHERE product_id = ? AND is_active = 1), 0) as reserved
    ''', (prod_id, prod_id)).fetchone()
    
    available = check['balance'] - check['reserved']
    
    if available < qty:
        conn.close()
        return jsonify({
            'error': f'Недостаточно товара для резерва! Доступно: {available} (остаток: {check["balance"]}, уже в резерве: {check["reserved"]})'
        }), 400

    cursor = conn.execute(
        "INSERT INTO reserves (product_id, quantity, user, comment, is_active) VALUES (?, ?, ?, ?, 1)",
        (prod_id, qty, user, comment)
    )
    new_id = cursor.lastrowid
    conn.commit()
    conn.close()
    
    return jsonify({
        'status': 'OK', 
        'message': f'Зарезервировано {qty} единиц. Пометка: {comment}',
        'reserve_id': new_id
    })


@app.route('/api/unreserve', methods=['POST'])
@login_required
def unreserve_product():
    username = session.get('username')
    
    if not can_reserve(username):
        return jsonify({'error': f'❌ У пользователя {username} нет прав на снятие резерва'}), 403

    data = request.get_json()
    reserve_id = data.get('reserve_id')
    prod_id = data.get('product_id')
    user = username

    if not reserve_id and not prod_id:
        return jsonify({'error': 'Не указан резерв или товар'}), 400

    conn = get_db()

    if reserve_id:
        reserve = conn.execute(
            "SELECT id, product_id, quantity, comment FROM reserves WHERE id = ? AND is_active = 1",
            (reserve_id,)
        ).fetchone()

        if not reserve:
            conn.close()
            return jsonify({'error': 'Резерв не найден или уже снят'}), 400

        conn.execute("UPDATE reserves SET is_active = 0 WHERE id = ?", (reserve_id,))
        conn.execute(
            "INSERT INTO stock_moves (product_id, quantity, user, comment) VALUES (?, ?, ?, ?)",
            (reserve['product_id'], 0, user, 
             f'Снят резерв #{reserve_id} ({reserve["quantity"]} шт, пометка: {reserve["comment"]})')
        )
        conn.commit()
        conn.close()
        return jsonify({
            'status': 'OK', 
            'message': f'Резерв #{reserve_id} снят ({reserve["quantity"]} шт)'
        })

    else:
        reserves = conn.execute(
            "SELECT id, quantity FROM reserves WHERE product_id = ? AND is_active = 1",
            (prod_id,)
        ).fetchall()

        if not reserves:
            conn.close()
            return jsonify({'error': 'Нет активных резервов для этого товара'}), 400

        total = sum(r['quantity'] for r in reserves)
        count = len(reserves)

        conn.execute(
            "UPDATE reserves SET is_active = 0 WHERE product_id = ? AND is_active = 1",
            (prod_id,)
        )
        conn.execute(
            "INSERT INTO stock_moves (product_id, quantity, user, comment) VALUES (?, ?, ?, ?)",
            (prod_id, 0, user, f'Сняты все резервы ({count} шт): {total} ед.')
        )
        conn.commit()
        conn.close()
        return jsonify({
            'status': 'OK', 
            'message': f'Снято {count} резервов на {total} единиц'
        })


@app.route('/api/reserves')
@login_required
def get_reserves():
    """Список активных резервов (время — пермское)"""
    conn = get_db()
    reserves = conn.execute('''
        SELECT 
            r.id,
            r.product_id,
            p.name as product_name,
            p.unit,
            r.quantity,
            r.user,
            r.comment,
            datetime(r.created_at, '+5 hours') as created_at
        FROM reserves r
        JOIN products p ON r.product_id = p.id
        WHERE r.is_active = 1
        ORDER BY p.name, r.created_at DESC
    ''').fetchall()
    conn.close()
    return jsonify([dict(row) for row in reserves])


# ============================================================
# ЭКСПОРТ В EXCEL
# ============================================================

@app.route('/export_excel')
@login_required
def export_excel():
    if session.get('role') != 'admin':
        return render_template('access_denied.html'), 403
    
    conn = get_db()
    data = conn.execute('''
        SELECT 
            p.id,
            p.name,
            COALESCE(SUM(sm.quantity), 0) as balance,
            COALESCE((
                SELECT SUM(r.quantity) 
                FROM reserves r 
                WHERE r.product_id = p.id AND r.is_active = 1
            ), 0) as reserved,
            p.unit
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id
        ORDER BY p.id
    ''').fetchall()
    conn.close()
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Остатки склада"
    
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
    
    headers = ['№ п/п', 'Наименование позиции', 'Ед. изм.', 'Остаток, шт', 'Резерв, шт', 'Доступно, шт']
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    for row_idx, item in enumerate(data, 2):
        available = item['balance'] - item['reserved']
        
        cell = ws.cell(row=row_idx, column=1, value=item['id'])
        cell.font = data_font
        cell.alignment = number_alignment
        cell.border = thin_border
        
        cell = ws.cell(row=row_idx, column=2, value=item['name'])
        cell.font = data_font
        cell.alignment = data_alignment
        cell.border = thin_border
        
        cell = ws.cell(row=row_idx, column=3, value=item['unit'] or 'шт')
        cell.font = data_font
        cell.alignment = number_alignment
        cell.border = thin_border
        
        cell = ws.cell(row=row_idx, column=4, value=item['balance'])
        cell.alignment = number_alignment
        cell.border = thin_border
        if item['balance'] < 5 and item['balance'] > 0:
            cell.font = Font(color="FF8F00", bold=True, size=10)
            cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        elif item['balance'] <= 0:
            cell.font = Font(color="FF0000", bold=True, size=10)
            cell.fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
        else:
            cell.font = Font(color="1B5E20", size=10)
        
        cell = ws.cell(row=row_idx, column=5, value=item['reserved'])
        cell.alignment = number_alignment
        cell.border = thin_border
        if item['reserved'] > 0:
            cell.font = Font(color="F57C00", bold=True, size=10)
            cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        else:
            cell.font = data_font
        
        cell = ws.cell(row=row_idx, column=6, value=available)
        cell.alignment = number_alignment
        cell.border = thin_border
        cell.font = Font(color="1B5E20", bold=True, size=10)
    
    column_widths = {'A': 8, 'B': 60, 'C': 12, 'D': 15, 'E': 15, 'F': 15}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    ws.freeze_panes = 'A2'
    
    total_items = len(data)
    low_items = len([item for item in data if 0 < item['balance'] < 5])
    zero_items = len([item for item in data if item['balance'] <= 0])
    reserved_items = len([item for item in data if item['reserved'] > 0])
    
    ws.append([])
    ws.append([f'Дата выгрузки: {get_perm_time().strftime("%d.%m.%Y %H:%M")} (Пермь)'])
    ws.append([f'Всего позиций: {total_items}'])
    ws.append([f'Позиций с остатком менее 5 шт: {low_items}'])
    ws.append([f'Позиций с нулевым остатком: {zero_items}'])
    ws.append([f'Позиций в резерве: {reserved_items}'])
    ws.append([f'Выгрузил: {session.get("username")} (администратор)'])
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Остатки_склада_{get_perm_time().strftime('%Y%m%d_%H%M')}.xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


# ============================================================
# ЗАПУСК
# ============================================================

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)