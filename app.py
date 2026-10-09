from flask import Flask, request, jsonify, render_template, session, redirect, url_for, send_file
import psycopg2
from psycopg2.extras import RealDictCursor
import bcrypt
from datetime import datetime, timedelta
import os
import io
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage
from functools import wraps
from dotenv import load_dotenv

# Загружаем переменные из .env
load_dotenv()

app = Flask(__name__)

# ============================================================
# БЕЗОПАСНОСТЬ
# ============================================================
SECRET_KEY = os.environ.get('SECRET_KEY')
if not SECRET_KEY:
    raise ValueError(
        "❌ Переменная окружения SECRET_KEY не задана!\n"
        "   Добавьте её в файл .env (локально) или в переменные окружения App Platform (на хостинге).\n"
        "   Сгенерировать можно командой: python -c \"import secrets; print(secrets.token_hex(32))\""
    )
app.secret_key = SECRET_KEY
app.permanent_session_lifetime = timedelta(hours=8)

# ============================================================
# ПОДКЛЮЧЕНИЕ К БД
# ============================================================
DATABASE_URL = os.environ.get('DATABASE_URL')
if not DATABASE_URL:
    raise ValueError(
        "❌ Переменная окружения DATABASE_URL не задана!\n"
        "   Добавьте её в файл .env (локально) или в переменные окружения App Platform (на хостинге)."
    )


# ============================================================
# ЦЕНЫ — КОНФИГУРАЦИЯ
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRICES_DIR = os.path.join(BASE_DIR, 'prices')
os.makedirs(PRICES_DIR, exist_ok=True)

PRICE_TABS = {
    'common':      'Общая цена',
    'partners':    'Цена партнеров',
    'armoservice': 'Армосервис',
}

ALLOWED_EXTENSIONS = {'.xlsx', '.xls'}


# ============================================================
# ПЕРМСКОЕ ВРЕМЯ (UTC+5)
# ============================================================
def get_perm_time():
    return datetime.utcnow() + timedelta(hours=2)


# ============================================================
# БАЗА ДАННЫХ
# ============================================================
def get_db():
    conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    return conn


# ============================================================
# БРЕНДЫ
# ============================================================
BRAND_RANGES = [
    ('Sondex',        1,   119),
    ('Alfa-Laval',    120, 216),
    ('Tranter',       217, 257),
    ('Funke РоСВЕП',  258, 314),
    ('Kelvion',       315, 362),
    ('Теплотекс-APV', 363, 369),
    ('Прочее',        370, 999999),
]


def get_special_brand_by_name(product_name):
    if not product_name:
        return None
    name_upper = product_name.upper().strip()
    if name_upper.startswith('SIGMA'):
        return 'SIGMA'
    name_normalized = name_upper.replace('О', 'O').replace('0', 'O')
    if (name_normalized.startswith('A4') or
        name_normalized.startswith('A6') or
        name_normalized.startswith('A8')):
        return 'ARES'
    if name_normalized.startswith('OO34') or name_normalized.startswith('O34'):
        return 'Теплотекс-APV'
    return None


def get_brand_by_position(position, product_name=None):
    special_brand = get_special_brand_by_name(product_name)
    if special_brand:
        return special_brand
    for brand_name, start, end in BRAND_RANGES:
        if start <= position <= end:
            return brand_name
    return 'Прочее'


BRAND_COLORS = {
    'Sondex':        '1a3c5e',
    'Alfa-Laval':    'B71C1C',
    'Tranter':       '4A148C',
    'Funke РоСВЕП':  '1B5E20',
    'Kelvion':       'E65100',
    'Теплотекс-APV': '006064',
    'SIGMA':         '6A1B9A',
    'ARES':          '00838F',
    'Прочее':        '424242',
}

BRAND_ORDER = {
    'Sondex':        1,
    'Alfa-Laval':    2,
    'Tranter':       3,
    'Funke РоСВЕП':  4,
    'Kelvion':       5,
    'Теплотекс-APV': 6,
    'SIGMA':         7,
    'ARES':          8,
    'Прочее':        9,
}


# ============================================================
# НОРМАЛИЗАЦИЯ ИМЁН
# ============================================================
_CYRILLIC_TO_LATIN = {
    'А': 'A', 'В': 'B', 'С': 'C', 'Е': 'E', 'Н': 'H',
    'К': 'K', 'М': 'M', 'О': 'O', 'Р': 'P', 'Т': 'T',
    'У': 'Y', 'Х': 'X', 'а': 'a', 'в': 'b', 'с': 'c',
    'е': 'e', 'н': 'h', 'к': 'k', 'м': 'm', 'о': 'o',
    'р': 'p', 'т': 't', 'у': 'y', 'х': 'x',
}


def normalize_product_name(name):
    if not name:
        return ''
    s = str(name).strip().upper()
    for cyr, lat in _CYRILLIC_TO_LATIN.items():
        s = s.replace(cyr, lat)
    s = re.sub(r'[«»"\'`.,;:!?()\[\]{}]', ' ', s)
    s = re.sub(r'[—–\-_/\\]', ' ', s)
    s = ' '.join(s.split())
    return s


def extract_excel_id(row):
    if not row:
        return None
    cell = row[0]
    if cell is None:
        return None
    if isinstance(cell, (int, float)):
        return int(cell)
    s = str(cell).strip()
    if s.isdigit():
        return int(s)
    return None


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


def prices_access_required(f):
    """
    Декоратор для страниц/API цен.
    Для API возвращает JSON 403, для страниц — HTML access_denied.html.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        username = session.get('username')

        if not username:
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Требуется авторизация'}), 401
            return redirect(url_for('login_page'))

        if not can_view_prices(username):
            if request.path.startswith('/api/'):
                return jsonify({'error': '❌ Доступ к ценам запрещён'}), 403
            return render_template('access_denied.html'), 403

        return f(*args, **kwargs)
    return decorated_function


# ============================================================
# ПРАВА ПОЛЬЗОВАТЕЛЕЙ
# ============================================================
CAN_SHIP_USERS = ['Павел', 'Валерий', 'Андрей']
CAN_RESERVE_USERS = ['Павел', 'Евгений', 'Виталий', 'Андрей']
CAN_RECEIVE_USERS = ['Павел', 'Валерий', 'Андрей']
CAN_EXPORT_EXCEL = ['Павел', 'Андрей', 'Евгений', 'Виталий']

# Кто видит цены (все, кроме Валерия)
CAN_VIEW_PRICES = ['Павел', 'Андрей', 'Евгений', 'Виталий']

# Кто может загружать/удалять файлы цен
CAN_MANAGE_PRICES = ['Павел', 'Андрей']


def can_ship(username):
    return username in CAN_SHIP_USERS


def can_reserve(username):
    return username in CAN_RESERVE_USERS


def can_receive(username):
    return username in CAN_RECEIVE_USERS


def can_export_excel(username):
    return username in CAN_EXPORT_EXCEL


def can_view_prices(username):
    return username in CAN_VIEW_PRICES


def can_manage_prices(username):
    return username in CAN_MANAGE_PRICES


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
    username = session.get('username')
    return render_template(
        'index.html',
        username=username,
        role=session.get('role'),
        can_view_prices=can_view_prices(username),
    )


@app.route('/history')
@login_required
def history():
    username = session.get('username')
    if not username:
        return render_template('access_denied.html'), 403

    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT 
            sm.id,
            p.name as product_name,
            sm.quantity,
            sm.user,
            sm.comment,
            to_char(sm.created_at + INTERVAL '5 hours', 'YYYY-MM-DD HH24:MI:SS') as created_at
        FROM stock_moves sm
        JOIN products p ON sm.product_id = p.id
        ORDER BY sm.created_at DESC
        LIMIT 200
    ''')
    logs = cur.fetchall()
    cur.close()
    conn.close()
    return render_template('history.html', logs=logs, username=session.get('username'))


@app.route('/reserves')
@login_required
def reserves_page():
    return render_template('reserves.html', username=session.get('username'))


@app.route('/admin/products')
@login_required
def admin_products():
    if session.get('role') != 'admin':
        return render_template('access_denied.html'), 403
    return render_template('admin_products.html', username=session.get('username'))


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login_page'))


# ============================================================
# ЦЕНЫ — СТРАНИЦА
# ============================================================
@app.route('/prices')
@login_required
@prices_access_required
def prices_page():
    username = session.get('username')
    can_manage = can_manage_prices(username)

    files_status = {}
    for key, name in PRICE_TABS.items():
        filename = f"{key}.xlsx"
        filepath = os.path.join(PRICES_DIR, filename)
        exists = os.path.exists(filepath)

        modified = None
        if exists:
            try:
                mtime = os.path.getmtime(filepath)
                modified = datetime.fromtimestamp(mtime).strftime('%d.%m.%Y %H:%M')
            except:
                pass

        files_status[key] = {
            'name': name,
            'exists': exists,
            'modified': modified,
            'filename': filename,
        }

    return render_template(
        'prices.html',
        username=username,
        can_manage=can_manage,
        files_status=files_status,
    )


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ЦЕН
# ============================================================
def load_prices_from_excel(tab_key):
    """
    Читает Excel с ценами.
    Возвращает:
      - header_row: список заголовков
      - data_rows: список dict-строк {id, name, unit, price, raw}
      - col_indexes: dict {id, name, unit, price}
    """
    filepath = os.path.join(PRICES_DIR, f"{tab_key}.xlsx")
    if not os.path.exists(filepath):
        return None, None, None

    wb = load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active

    raw_rows = []
    for row in ws.iter_rows(values_only=True):
        if all(c is None or str(c).strip() == '' for c in row):
            continue
        raw_rows.append(list(row))
    wb.close()

    if not raw_rows:
        return None, [], None

    # Ищем заголовок
    header_idx = 0
    col_indexes = {'id': 0, 'name': None, 'unit': None, 'price': None}

    for i, row in enumerate(raw_rows[:50]):
        has_name = False
        for j, cell in enumerate(row):
            s = str(cell).strip().lower() if cell is not None else ''
            if 'наименование' in s:
                col_indexes['name'] = j
                has_name = True
                header_idx = i
            elif 'ед' in s and ('изм' in s or s == 'ед.'):
                col_indexes['unit'] = j
            elif 'цена' in s or 'price' in s:
                col_indexes['price'] = j
        if has_name:
            break

    # Дефолты
    if col_indexes['name'] is None:
        col_indexes['name'] = 1
        col_indexes['unit'] = 2
        col_indexes['price'] = 6
        header_idx = -1

    header_row = raw_rows[header_idx] if header_idx >= 0 else ['№', 'Наименование', 'Ед. изм.', '', '', '', 'Цена']

    data_rows = []
    for row in raw_rows[header_idx + 1:]:
        if len(row) <= col_indexes['name']:
            continue

        name_val = row[col_indexes['name']]
        name_str = str(name_val).strip() if name_val is not None else ''

        if not name_str:
            continue

        # Пропускаем заголовки брендов и служебные строки
        if re.search(r'\(\d+\s*позиц', name_str, re.IGNORECASE):
            continue
        if name_str.startswith(('Всего', 'Дата', 'Позиций', 'Выгрузил')):
            continue

        # Цена
        price_val = None
        if col_indexes['price'] is not None and len(row) > col_indexes['price']:
            p = row[col_indexes['price']]
            if isinstance(p, (int, float)):
                price_val = float(p)
            elif p is not None and str(p).strip() != '':
                try:
                    price_val = float(str(p).replace(' ', '').replace(',', '.'))
                except (ValueError, TypeError):
                    price_val = None

        # Ед. изм.
        unit_val = ''
        if col_indexes['unit'] is not None and len(row) > col_indexes['unit']:
            u = row[col_indexes['unit']]
            unit_val = str(u).strip() if u is not None else ''

        data_rows.append({
            'id': extract_excel_id(row),
            'name': name_str,
            'unit': unit_val,
            'price': price_val,
            'raw': list(row),
        })

    return header_row, data_rows, col_indexes


def get_db_balances():
    """
    Возвращает:
      by_id: {product_id: {id, name, unit, balance, reserved, available, sort_order}}
      by_name: {normalized_name: {...}}
    """
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT
            p.id,
            p.name,
            p.unit,
            p.sort_order,
            COALESCE(SUM(sm.quantity), 0) AS balance,
            COALESCE((
                SELECT SUM(r.quantity)
                FROM reserves r
                WHERE r.product_id = p.id AND r.is_active = 1
            ), 0) AS reserved
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id, p.name, p.unit, p.sort_order
        ORDER BY p.sort_order, p.id
    ''')
    rows = cur.fetchall()
    cur.close()
    conn.close()

    by_id = {}
    by_name = {}
    for r in rows:
        item = dict(r)
        item['available'] = item['balance'] - item['reserved']
        try:
            item['_sort_order'] = float(item['sort_order']) if item['sort_order'] is not None else 0
        except (ValueError, TypeError):
            item['_sort_order'] = 0
        by_id[item['id']] = item
        by_name[normalize_product_name(item['name'])] = item

    return by_id, by_name


# ============================================================
# ЦЕНЫ — API
# ============================================================
@app.route('/api/prices/upload', methods=['POST'])
@login_required
@prices_access_required
def prices_upload():
    username = session.get('username')

    if not can_manage_prices(username):
        return jsonify({'error': '❌ Только администраторы могут загружать файлы'}), 403

    tab_key = request.form.get('tab')
    if tab_key not in PRICE_TABS:
        return jsonify({'error': '❌ Неверная вкладка'}), 400

    if 'file' not in request.files:
        return jsonify({'error': '❌ Файл не выбран'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': '❌ Файл не выбран'}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({'error': '❌ Разрешены только файлы .xlsx и .xls'}), 400

    save_path = os.path.join(PRICES_DIR, f"{tab_key}.xlsx")
    try:
        file.save(save_path)
        print(f"✅ Файл цен загружен: {save_path}")
    except Exception as e:
        return jsonify({'error': f'❌ Ошибка сохранения: {str(e)}'}), 500

    return jsonify({
        'status': 'OK',
        'message': f'Файл для вкладки «{PRICE_TABS[tab_key]}» загружен'
    })


@app.route('/api/prices/data/<tab_key>')
@login_required
@prices_access_required
def prices_data(tab_key):
    """
    Возвращает таблицу:
    - наименование, ед. изм., цена — из Excel
    - остаток, резерв, доступно — из БД
    """
    if tab_key not in PRICE_TABS:
        return jsonify({'error': 'Неверная вкладка'}), 400

    filepath = os.path.join(PRICES_DIR, f"{tab_key}.xlsx")
    if not os.path.exists(filepath):
        return jsonify({'exists': False})

    header_row, excel_rows, col_indexes = load_prices_from_excel(tab_key)
    if header_row is None:
        return jsonify({'exists': True, 'rows': [], 'empty': True})

    db_by_id, db_by_name = get_db_balances()

    result_rows = [
        ['№ п/п', 'Наименование позиции', 'Ед. изм.', 'Остаток, шт', 'Резерв, шт', 'Доступно, шт', 'Цена, руб.']
    ]

    matched_by_id = 0
    matched_by_name = 0
    unmatched = []
    counter = 0

    for ex in excel_rows:
        db_item = None

        if ex['id'] and ex['id'] in db_by_id:
            db_item = db_by_id[ex['id']]
            matched_by_id += 1
        else:
            norm = normalize_product_name(ex['name'])
            if norm in db_by_name:
                db_item = db_by_name[norm]
                matched_by_name += 1

        counter += 1

        if db_item:
            result_rows.append([
                counter,
                ex['name'],
                ex['unit'] or db_item.get('unit') or 'шт',
                db_item['balance'],
                db_item['reserved'],
                db_item['available'],
                ex['price'] if ex['price'] is not None else '',
            ])
        else:
            unmatched.append({'excel_id': ex['id'], 'name': ex['name']})
            result_rows.append([
                counter,
                ex['name'],
                ex['unit'] or 'шт',
                '—',
                '—',
                '—',
                ex['price'] if ex['price'] is not None else '',
            ])

    mtime = os.path.getmtime(filepath)
    modified = datetime.fromtimestamp(mtime).strftime('%d.%m.%Y %H:%M:%S')

    return jsonify({
        'exists': True,
        'modified': modified,
        'rows': result_rows,
        'match_stats': {
            'total_in_db': len(db_by_id),
            'total_in_excel': len(excel_rows),
            'matched_by_id': matched_by_id,
            'matched_by_name': matched_by_name,
            'unmatched_count': len(unmatched),
        },
        'unmatched': unmatched[:50],
    })


@app.route('/api/prices/delete/<tab_key>', methods=['POST'])
@login_required
@prices_access_required
def prices_delete(tab_key):
    username = session.get('username')

    if not can_manage_prices(username):
        return jsonify({'error': '❌ Только администраторы могут удалять файлы'}), 403

    if tab_key not in PRICE_TABS:
        return jsonify({'error': '❌ Неверная вкладка'}), 400

    filepath = os.path.join(PRICES_DIR, f"{tab_key}.xlsx")
    if not os.path.exists(filepath):
        return jsonify({'error': '❌ Файл не найден'}), 404

    try:
        os.remove(filepath)
        print(f"🗑️ Файл цен удалён: {filepath}")
    except Exception as e:
        return jsonify({'error': f'❌ Ошибка удаления: {str(e)}'}), 500

    return jsonify({
        'status': 'OK',
        'message': f'Файл для вкладки «{PRICE_TABS[tab_key]}» удалён'
    })


@app.route('/api/prices/download/<tab_key>')
@login_required
@prices_access_required
def prices_download(tab_key):
    """
    Генерирует Excel в том же стиле, что и /export_excel.
    Доступ — только для CAN_VIEW_PRICES.
    """
    if tab_key not in PRICE_TABS:
        return jsonify({'error': 'Неверная вкладка'}), 400

    filepath = os.path.join(PRICES_DIR, f"{tab_key}.xlsx")
    if not os.path.exists(filepath):
        return jsonify({'error': 'Файл с ценами не загружен'}), 404

    header_row, excel_rows, col_indexes = load_prices_from_excel(tab_key)
    if header_row is None:
        return jsonify({'error': 'Не удалось прочитать файл'}), 500

    db_by_id, db_by_name = get_db_balances()
    username = session.get('username')

    # ---------- 1. Обогащаем строки Excel данными из БД ----------
    enriched_products = []
    matched_by_id = 0
    matched_by_name = 0
    unmatched_names = []

    for ex in excel_rows:
        db_item = None
        match_type = None

        if ex['id'] and ex['id'] in db_by_id:
            db_item = db_by_id[ex['id']]
            match_type = 'id'
            matched_by_id += 1
        else:
            norm = normalize_product_name(ex['name'])
            if norm in db_by_name:
                db_item = db_by_name[norm]
                match_type = 'name'
                matched_by_name += 1

        if db_item:
            enriched_products.append({
                'id': db_item['id'],
                'name': ex['name'] or db_item['name'],
                'unit': ex['unit'] or db_item.get('unit') or 'шт',
                'balance': db_item['balance'],
                'reserved': db_item['reserved'],
                'available': db_item['available'],
                'price': ex['price'],
                'sort_order': db_item.get('_sort_order', 0),
                'matched': True,
                'match_type': match_type,
            })
        else:
            unmatched_names.append(ex['name'])
            enriched_products.append({
                'id': ex['id'] or 0,
                'name': ex['name'],
                'unit': ex['unit'] or 'шт',
                'balance': None,
                'reserved': None,
                'available': None,
                'price': ex['price'],
                'sort_order': 999999,
                'matched': False,
                'match_type': None,
            })

    # ---------- 2. Сортируем по sort_order из БД ----------
    enriched_products.sort(key=lambda p: (p['sort_order'], p['id']))

    # ---------- 3. Группируем по брендам ----------
    brands_data = {}
    for prod in enriched_products:
        if prod['matched']:
            brand_name = get_brand_by_position(prod['id'], prod['name'])
        else:
            brand_name = 'Прочее'

        if brand_name not in brands_data:
            brands_data[brand_name] = []
        brands_data[brand_name].append(prod)

    sorted_brands = sorted(brands_data.items(), key=lambda x: BRAND_ORDER.get(x[0], 99))

    # ---------- 4. Создаём Excel ----------
    wb = Workbook()
    ws = wb.active
    ws.title = PRICE_TABS[tab_key][:30]

    column_widths = {
        'A': 8, 'B': 60, 'C': 12,
        'D': 15, 'E': 15, 'F': 15,
        'G': 18,
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    static_dir = os.path.join(BASE_DIR, 'static')
    logo_path = os.path.join(static_dir, 'logo.png')

    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 900
            img.height = 115
            ws.add_image(img, 'A1')
        except Exception as e:
            print(f"⚠️ Не удалось добавить логотип: {e}")

    ws.row_dimensions[1].height = 29
    ws.row_dimensions[2].height = 29
    ws.row_dimensions[3].height = 29
    ws.row_dimensions[4].height = 29
    ws.row_dimensions[5].height = 8
    ws.row_dimensions[6].height = 5

    red_fill = PatternFill(start_color="E53935", end_color="E53935", fill_type="solid")
    blue_fill = PatternFill(start_color="1a3c5e", end_color="1a3c5e", fill_type="solid")
    for col in range(1, 8):
        ws.cell(row=5, column=col).fill = red_fill
        ws.cell(row=6, column=col).fill = blue_fill

    ws.row_dimensions[7].height = 10

    HEADER_ROW = 8
    current_row = HEADER_ROW + 1

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1a3c5e", end_color="1a3c5e", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    data_font = Font(size=10)
    data_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    number_alignment = Alignment(horizontal="center", vertical="center")
    price_alignment = Alignment(horizontal="right", vertical="center")

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    headers = [
        '№ п/п',
        'Наименование позиции',
        'Ед. изм.',
        'Остаток, шт',
        'Резерв, шт',
        'Доступно, шт',
        'Цена, руб.'
    ]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    ws.row_dimensions[HEADER_ROW].height = 25

    counter = 0
    brand_header_font = Font(bold=True, color="FFFFFF", size=12)
    brand_alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for brand_name, products in sorted_brands:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)

        brand_color = BRAND_COLORS.get(brand_name, '1a3c5e')
        brand_fill = PatternFill(start_color=brand_color, end_color=brand_color, fill_type="solid")

        header_cell = ws.cell(row=current_row, column=1, value=f"  {brand_name}  ({len(products)} позиций)")
        header_cell.font = brand_header_font
        header_cell.fill = brand_fill
        header_cell.alignment = brand_alignment

        ws.row_dimensions[current_row].height = 22
        current_row += 1

        for prod in products:
            counter += 1

            cell = ws.cell(row=current_row, column=1, value=counter)
            cell.font = data_font
            cell.alignment = number_alignment
            cell.border = thin_border

            cell = ws.cell(row=current_row, column=2, value=prod['name'])
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border

            cell = ws.cell(row=current_row, column=3, value=prod['unit'] or 'шт')
            cell.font = data_font
            cell.alignment = number_alignment
            cell.border = thin_border

            if prod['balance'] is not None:
                cell = ws.cell(row=current_row, column=4, value=prod['balance'])
                cell.alignment = number_alignment
                cell.border = thin_border
                if prod['balance'] < 5 and prod['balance'] > 0:
                    cell.font = Font(color="FF8F00", bold=True, size=10)
                    cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
                elif prod['balance'] <= 0:
                    cell.font = Font(color="FF0000", bold=True, size=10)
                    cell.fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
                else:
                    cell.font = Font(color="1B5E20", size=10)
            else:
                cell = ws.cell(row=current_row, column=4, value='—')
                cell.font = Font(color="90A4AE", italic=True, size=10)
                cell.alignment = number_alignment
                cell.border = thin_border
                cell.fill = PatternFill(start_color="ECEFF1", end_color="ECEFF1", fill_type="solid")

            if prod['reserved'] is not None:
                cell = ws.cell(row=current_row, column=5, value=prod['reserved'])
                cell.alignment = number_alignment
                cell.border = thin_border
                if prod['reserved'] > 0:
                    cell.font = Font(color="F57C00", bold=True, size=10)
                    cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
                else:
                    cell.font = data_font
            else:
                cell = ws.cell(row=current_row, column=5, value='—')
                cell.font = Font(color="90A4AE", italic=True, size=10)
                cell.alignment = number_alignment
                cell.border = thin_border
                cell.fill = PatternFill(start_color="ECEFF1", end_color="ECEFF1", fill_type="solid")

            if prod['available'] is not None:
                cell = ws.cell(row=current_row, column=6, value=prod['available'])
                cell.alignment = number_alignment
                cell.border = thin_border
                if prod['available'] <= 0:
                    cell.font = Font(color="FF0000", bold=True, size=10)
                else:
                    cell.font = Font(color="1B5E20", bold=True, size=10)
            else:
                cell = ws.cell(row=current_row, column=6, value='—')
                cell.font = Font(color="90A4AE", italic=True, size=10)
                cell.alignment = number_alignment
                cell.border = thin_border
                cell.fill = PatternFill(start_color="ECEFF1", end_color="ECEFF1", fill_type="solid")

            if prod['price'] is not None:
                cell = ws.cell(row=current_row, column=7, value=prod['price'])
                cell.alignment = price_alignment
                cell.border = thin_border
                cell.number_format = '#,##0.00'
                cell.font = Font(color="1B5E20", bold=True, size=10)
                cell.fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
            else:
                cell = ws.cell(row=current_row, column=7, value='—')
                cell.alignment = price_alignment
                cell.border = thin_border
                cell.font = Font(color="90A4AE", italic=True, size=10)

            current_row += 1

        ws.row_dimensions[current_row].height = 8
        current_row += 1

    ws.freeze_panes = ws.cell(row=HEADER_ROW + 1, column=1)

    last_row = current_row + 1

    total_items = len(enriched_products)
    low_items = len([p for p in enriched_products if p['balance'] is not None and 0 < p['balance'] < 5])
    zero_items = len([p for p in enriched_products if p['balance'] is not None and p['balance'] <= 0])
    reserved_items = len([p for p in enriched_products if p['reserved'] is not None and p['reserved'] > 0])
    with_price = len([p for p in enriched_products if p['price'] is not None])

    info_font = Font(italic=True, size=9, color="777777")

    ws.cell(row=last_row, column=2,
            value=f'Дата выгрузки: {get_perm_time().strftime("%d.%m.%Y %H:%M")} (Пермь)').font = info_font
    ws.cell(row=last_row + 1, column=2,
            value=f'Всего позиций: {total_items}').font = info_font
    ws.cell(row=last_row + 2, column=2,
            value=f'Позиций с ценой: {with_price}').font = info_font
    ws.cell(row=last_row + 3, column=2,
            value=f'Позиций с остатком менее 5 шт: {low_items}').font = info_font
    ws.cell(row=last_row + 4, column=2,
            value=f'Позиций с нулевым остатком: {zero_items}').font = info_font
    ws.cell(row=last_row + 5, column=2,
            value=f'Позиций в резерве: {reserved_items}').font = info_font
    ws.cell(row=last_row + 6, column=2,
            value=f'Сопоставлено с БД: по ID — {matched_by_id}, по имени — {matched_by_name}').font = info_font

    if unmatched_names:
        ws.cell(row=last_row + 7, column=2,
                value=f'Не найдено в БД: {len(unmatched_names)}').font = Font(italic=True, size=9, color="D32F2F")

    ws.cell(row=last_row + 8, column=2,
            value=f'Выгрузил: {username}').font = info_font

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"{PRICE_TABS[tab_key]}_{get_perm_time().strftime('%Y%m%d_%H%M')}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


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
    cur = conn.cursor()
    cur.execute("SELECT id, name, password_hash, role FROM users WHERE name = %s", (username,))
    user = cur.fetchone()
    cur.close()
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
    role = session.get('role')
    return jsonify({
        'username': username,
        'role': role,
        'can_ship': can_ship(username),
        'can_reserve': can_reserve(username),
        'can_receive': can_receive(username),
        'can_view_history': True,
        'can_export_excel': can_export_excel(username),
        'can_view_prices': can_view_prices(username),
        'can_manage_prices': can_manage_prices(username),
    })


@app.route('/api/users')
@login_required
def get_users():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT name FROM users ORDER BY name")
    users = cur.fetchall()
    cur.close()
    conn.close()
    return jsonify([u['name'] for u in users])


# ============================================================
# API - ОСТАТКИ
# ============================================================
@app.route('/api/balances')
@login_required
def balances():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
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
        GROUP BY p.id, p.name, p.unit, p.sort_order
        ORDER BY p.sort_order, p.id
    ''')
    data = cur.fetchall()
    cur.close()
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
    username = session.get('username')

    if not can_ship(username):
        return jsonify({'error': f'❌ У пользователя {username} нет прав на отгрузку'}), 403

    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    comment = data.get('comment', 'Отгрузка')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    cur = conn.cursor()

    cur.execute('''
        SELECT 
            COALESCE((SELECT SUM(quantity) FROM stock_moves WHERE product_id = %s), 0) as balance,
            COALESCE((SELECT SUM(quantity) FROM reserves WHERE product_id = %s AND is_active = 1), 0) as reserved
    ''', (prod_id, prod_id))
    check = cur.fetchone()

    balance = check['balance']
    reserved = check['reserved']
    available = balance - reserved

    if qty > available:
        cur.close()
        conn.close()
        if reserved > 0:
            return jsonify({
                'error': f'❌ Нельзя отгрузить {qty} шт! Доступно: {available} (остаток: {balance}, в резерве: {reserved})'
            }), 400
        else:
            return jsonify({
                'error': f'❌ Недостаточно товара! Остаток: {balance}'
            }), 400

    cur.execute(
        "INSERT INTO stock_moves (product_id, quantity, \"user\", comment) VALUES (%s, %s, %s, %s)",
        (prod_id, -abs(qty), username, comment)
    )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'status': 'OK', 'message': f'Отгружено {qty} единиц'})


@app.route('/api/receive', methods=['POST'])
@login_required
def receive():
    username = session.get('username')

    if not can_receive(username):
        return jsonify({'error': f'❌ У пользователя {username} нет прав на пополнение склада'}), 403

    data = request.get_json()
    prod_id = data.get('product_id')
    qty = data.get('quantity')
    comment = data.get('comment', 'Пополнение')

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO stock_moves (product_id, quantity, \"user\", comment) VALUES (%s, %s, %s, %s)",
        (prod_id, abs(qty), username, comment)
    )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({'status': 'OK', 'message': f'Добавлено {qty} единиц'})


# ============================================================
# API - УПРАВЛЕНИЕ ТОВАРАМИ
# ============================================================
@app.route('/api/admin/products', methods=['GET'])
@login_required
@role_required(['admin'])
def admin_get_products():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT 
            p.id, 
            p.name, 
            p.unit, 
            p.sort_order,
            COALESCE(SUM(sm.quantity), 0) as balance,
            COALESCE((
                SELECT SUM(r.quantity) 
                FROM reserves r 
                WHERE r.product_id = p.id AND r.is_active = 1
            ), 0) as reserved,
            (SELECT COUNT(*) FROM stock_moves WHERE product_id = p.id) as moves_count
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id, p.name, p.unit, p.sort_order
        ORDER BY p.sort_order, p.id
    ''')
    data = cur.fetchall()
    cur.close()
    conn.close()

    result = []
    for row in data:
        item = dict(row)
        item['available'] = item['balance'] - item['reserved']
        result.append(item)

    return jsonify(result)


@app.route('/api/admin/products', methods=['POST'])
@login_required
@role_required(['admin'])
def admin_add_product():
    data = request.get_json()
    name = (data.get('name') or '').strip()
    unit = (data.get('unit') or 'шт').strip() or 'шт'
    initial_qty = data.get('initial_qty', 0)
    after_id = data.get('after_id')
    user = session.get('username', 'admin')

    if not name:
        return jsonify({'error': 'Укажите наименование товара'}), 400

    try:
        initial_qty = int(initial_qty)
        if initial_qty < 0:
            initial_qty = 0
    except (ValueError, TypeError):
        initial_qty = 0

    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT id FROM products WHERE name = %s", (name,))
    existing = cur.fetchone()
    if existing:
        cur.close()
        conn.close()
        return jsonify({'error': f'Товар с именем "{name}" уже существует'}), 400

    new_sort_order = None

    if after_id:
        try:
            after_id = int(after_id)
            cur.execute("SELECT sort_order FROM products WHERE id = %s", (after_id,))
            after_product = cur.fetchone()

            if after_product:
                after_order = float(after_product['sort_order']) if after_product['sort_order'] is not None else 0

                cur.execute(
                    "SELECT MIN(sort_order) as next_order FROM products WHERE sort_order > %s",
                    (after_order,)
                )
                next_product = cur.fetchone()
                next_order = float(next_product['next_order']) if next_product['next_order'] is not None else None

                if next_order is not None:
                    new_sort_order = (after_order + next_order) / 2
                else:
                    new_sort_order = after_order + 1
            else:
                cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 as new_order FROM products")
                new_sort_order = float(cur.fetchone()['new_order'])
        except (ValueError, TypeError):
            cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 as new_order FROM products")
            new_sort_order = float(cur.fetchone()['new_order'])
    else:
        cur.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 as new_order FROM products")
        new_sort_order = float(cur.fetchone()['new_order'])

    cur.execute(
        "INSERT INTO products (name, unit, sort_order) VALUES (%s, %s, %s) RETURNING id",
        (name, unit, new_sort_order)
    )
    new_id = cur.fetchone()['id']

    if initial_qty > 0:
        cur.execute(
            "INSERT INTO stock_moves (product_id, quantity, \"user\", comment) VALUES (%s, %s, %s, %s)",
            (new_id, initial_qty, user, 'Начальный остаток')
        )

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({
        'status': 'OK',
        'message': f'Товар "{name}" добавлен (остаток: {initial_qty} {unit})',
        'product_id': new_id,
        'sort_order': new_sort_order
    })


@app.route('/api/admin/products/<int:product_id>', methods=['PUT'])
@login_required
@role_required(['admin'])
def admin_update_product(product_id):
    data = request.get_json()
    name = (data.get('name') or '').strip()
    unit = (data.get('unit') or 'шт').strip() or 'шт'
    sort_order = data.get('sort_order')

    if not name:
        return jsonify({'error': 'Укажите наименование товара'}), 400

    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT id FROM products WHERE id = %s", (product_id,))
    product = cur.fetchone()
    if not product:
        cur.close()
        conn.close()
        return jsonify({'error': 'Товар не найден'}), 404

    cur.execute("SELECT id FROM products WHERE name = %s AND id != %s", (name, product_id))
    duplicate = cur.fetchone()
    if duplicate:
        cur.close()
        conn.close()
        return jsonify({'error': f'Товар с именем "{name}" уже существует'}), 400

    if sort_order is not None and sort_order != '':
        try:
            sort_order = float(sort_order)
            cur.execute(
                "UPDATE products SET name = %s, unit = %s, sort_order = %s WHERE id = %s",
                (name, unit, sort_order, product_id)
            )
        except (ValueError, TypeError):
            cur.execute(
                "UPDATE products SET name = %s, unit = %s WHERE id = %s",
                (name, unit, product_id)
            )
    else:
        cur.execute(
            "UPDATE products SET name = %s, unit = %s WHERE id = %s",
            (name, unit, product_id)
        )

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({'status': 'OK', 'message': f'Товар обновлён'})


@app.route('/api/admin/products/<int:product_id>', methods=['DELETE'])
@login_required
@role_required(['admin'])
def admin_delete_product(product_id):
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT id, name FROM products WHERE id = %s", (product_id,))
    product = cur.fetchone()
    if not product:
        cur.close()
        conn.close()
        return jsonify({'error': 'Товар не найден'}), 404

    cur.execute("SELECT COUNT(*) as cnt FROM stock_moves WHERE product_id = %s", (product_id,))
    moves = cur.fetchone()

    cur.execute("SELECT COUNT(*) as cnt FROM reserves WHERE product_id = %s AND is_active = 1", (product_id,))
    reserves = cur.fetchone()

    if moves['cnt'] > 1 or reserves['cnt'] > 0:
        cur.close()
        conn.close()
        reasons = []
        if moves['cnt'] > 1:
            reasons.append(f'движений: {moves["cnt"]}')
        if reserves['cnt'] > 0:
            reasons.append(f'активных резервов: {reserves["cnt"]}')
        return jsonify({
            'error': f'Нельзя удалить товар "{product["name"]}" — есть история: {", ".join(reasons)}.'
        }), 400

    cur.execute("DELETE FROM stock_moves WHERE product_id = %s", (product_id,))
    cur.execute("DELETE FROM reserves WHERE product_id = %s", (product_id,))
    cur.execute("DELETE FROM products WHERE id = %s", (product_id,))
    conn.commit()
    cur.close()
    conn.close()

    return jsonify({'status': 'OK', 'message': f'Товар "{product["name"]}" удалён'})


# ============================================================
# API - РЕЗЕРВЫ
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
    comment = (data.get('comment') or 'Без пометки').strip() or 'Без пометки'

    if not prod_id or not qty or qty <= 0:
        return jsonify({'error': 'Некорректные данные'}), 400

    conn = get_db()
    cur = conn.cursor()

    cur.execute('''
        SELECT 
            COALESCE((SELECT SUM(quantity) FROM stock_moves WHERE product_id = %s), 0) as balance,
            COALESCE((SELECT SUM(quantity) FROM reserves WHERE product_id = %s AND is_active = 1), 0) as reserved
    ''', (prod_id, prod_id))
    check = cur.fetchone()

    available = check['balance'] - check['reserved']

    if available < qty:
        cur.close()
        conn.close()
        return jsonify({
            'error': f'Недостаточно товара для резерва! Доступно: {available}'
        }), 400

    cur.execute(
        "INSERT INTO reserves (product_id, quantity, \"user\", comment, is_active) VALUES (%s, %s, %s, %s, 1) RETURNING id",
        (prod_id, qty, username, comment)
    )
    new_id = cur.fetchone()['id']
    conn.commit()
    cur.close()
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

    if not reserve_id and not prod_id:
        return jsonify({'error': 'Не указан резерв или товар'}), 400

    conn = get_db()
    cur = conn.cursor()

    if reserve_id:
        cur.execute(
            "SELECT id, product_id, quantity, comment FROM reserves WHERE id = %s AND is_active = 1",
            (reserve_id,)
        )
        reserve = cur.fetchone()

        if not reserve:
            cur.close()
            conn.close()
            return jsonify({'error': 'Резерв не найден или уже снят'}), 400

        cur.execute("UPDATE reserves SET is_active = 0 WHERE id = %s", (reserve_id,))
        cur.execute(
            "INSERT INTO stock_moves (product_id, quantity, \"user\", comment) VALUES (%s, %s, %s, %s)",
            (reserve['product_id'], 0, username,
             f'Снят резерв #{reserve_id} ({reserve["quantity"]} шт, пометка: {reserve["comment"]})')
        )
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({
            'status': 'OK',
            'message': f'Резерв #{reserve_id} снят ({reserve["quantity"]} шт)'
        })

    else:
        cur.execute(
            "SELECT id, quantity FROM reserves WHERE product_id = %s AND is_active = 1",
            (prod_id,)
        )
        reserves = cur.fetchall()

        if not reserves:
            cur.close()
            conn.close()
            return jsonify({'error': 'Нет активных резервов для этого товара'}), 400

        total = sum(r['quantity'] for r in reserves)
        count = len(reserves)

        cur.execute(
            "UPDATE reserves SET is_active = 0 WHERE product_id = %s AND is_active = 1",
            (prod_id,)
        )
        cur.execute(
            "INSERT INTO stock_moves (product_id, quantity, \"user\", comment) VALUES (%s, %s, %s, %s)",
            (prod_id, 0, username, f'Сняты все резервы ({count} шт): {total} ед.')
        )
        conn.commit()
        cur.close()
        conn.close()
        return jsonify({
            'status': 'OK',
            'message': f'Снято {count} резервов на {total} единиц'
        })


@app.route('/api/reserves')
@login_required
def get_reserves():
    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT 
            r.id,
            r.product_id,
            p.name as product_name,
            p.unit,
            r.quantity,
            r.user,
            r.comment,
            to_char(r.created_at + INTERVAL '5 hours', 'YYYY-MM-DD HH24:MI:SS') as created_at
        FROM reserves r
        JOIN products p ON r.product_id = p.id
        WHERE r.is_active = 1
        ORDER BY p.sort_order, p.id, r.created_at DESC
    ''')
    reserves = cur.fetchall()
    cur.close()
    conn.close()
    return jsonify([dict(row) for row in reserves])


# ============================================================
# ЭКСПОРТ В EXCEL (главный склад)
# ============================================================
@app.route('/export_excel')
@login_required
def export_excel():
    username = session.get('username')

    if not can_export_excel(username):
        return render_template('access_denied.html'), 403

    conn = get_db()
    cur = conn.cursor()
    cur.execute('''
        SELECT 
            p.id,
            p.name,
            p.sort_order,
            COALESCE(SUM(sm.quantity), 0) as balance,
            COALESCE((
                SELECT SUM(r.quantity) 
                FROM reserves r 
                WHERE r.product_id = p.id AND r.is_active = 1
            ), 0) as reserved,
            p.unit
        FROM products p
        LEFT JOIN stock_moves sm ON p.id = sm.product_id
        GROUP BY p.id, p.name, p.unit, p.sort_order
        ORDER BY p.sort_order, p.id
    ''')
    data = cur.fetchall()
    cur.close()
    conn.close()

    all_products = []
    for item in data:
        item_dict = dict(item)
        item_dict['available'] = item_dict['balance'] - item_dict['reserved']

        try:
            item_dict['_sort_order'] = float(item_dict.get('sort_order', item_dict['id'])) if item_dict.get('sort_order') is not None else 0
        except (ValueError, TypeError):
            item_dict['_sort_order'] = 0

        all_products.append(item_dict)

    all_products.sort(key=lambda x: (x['_sort_order'], x['id']))

    brands_data = {}
    for idx, item_dict in enumerate(all_products, start=1):
        brand_name = get_brand_by_position(idx, item_dict['name'])
        item_dict['_brand_name'] = brand_name

        if brand_name not in brands_data:
            brands_data[brand_name] = []
        brands_data[brand_name].append(item_dict)

    sorted_brands = sorted(brands_data.items(), key=lambda x: BRAND_ORDER.get(x[0], 99))

    wb = Workbook()
    ws = wb.active
    ws.title = "Остатки склада"

    column_widths = {
        'A': 8, 'B': 60, 'C': 12,
        'D': 15, 'E': 15, 'F': 15,
        'G': 18,
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    static_dir = os.path.join(BASE_DIR, 'static')
    logo_path = os.path.join(static_dir, 'logo.png')

    if os.path.exists(logo_path):
        try:
            img = XLImage(logo_path)
            img.width = 900
            img.height = 115
            ws.add_image(img, 'A1')
        except Exception as e:
            print(f"⚠️ Не удалось добавить логотип: {e}")

    ws.row_dimensions[1].height = 29
    ws.row_dimensions[2].height = 29
    ws.row_dimensions[3].height = 29
    ws.row_dimensions[4].height = 29
    ws.row_dimensions[5].height = 8
    ws.row_dimensions[6].height = 5

    red_fill = PatternFill(start_color="E53935", end_color="E53935", fill_type="solid")
    blue_fill = PatternFill(start_color="1a3c5e", end_color="1a3c5e", fill_type="solid")
    for col in range(1, 8):
        ws.cell(row=5, column=col).fill = red_fill
        ws.cell(row=6, column=col).fill = blue_fill

    ws.row_dimensions[7].height = 10

    HEADER_ROW = 8
    current_row = HEADER_ROW + 1

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

    headers = ['№ п/п', 'Наименование позиции', 'Ед. изм.', 'Остаток, шт', 'Резерв, шт', 'Доступно, шт', 'Цена, руб.']

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=HEADER_ROW, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    ws.row_dimensions[HEADER_ROW].height = 25

    counter = 0
    brand_header_font = Font(bold=True, color="FFFFFF", size=12)
    brand_alignment = Alignment(horizontal="left", vertical="center", indent=1)

    for brand_name, products in sorted_brands:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)

        brand_color = BRAND_COLORS.get(brand_name, '1a3c5e')
        brand_fill = PatternFill(start_color=brand_color, end_color=brand_color, fill_type="solid")

        header_cell = ws.cell(row=current_row, column=1, value=f"  {brand_name}  ({len(products)} позиций)")
        header_cell.font = brand_header_font
        header_cell.fill = brand_fill
        header_cell.alignment = brand_alignment

        ws.row_dimensions[current_row].height = 22
        current_row += 1

        for item in products:
            counter += 1

            cell = ws.cell(row=current_row, column=1, value=counter)
            cell.font = data_font
            cell.alignment = number_alignment
            cell.border = thin_border

            cell = ws.cell(row=current_row, column=2, value=item['name'])
            cell.font = data_font
            cell.alignment = data_alignment
            cell.border = thin_border

            cell = ws.cell(row=current_row, column=3, value=item['unit'] or 'шт')
            cell.font = data_font
            cell.alignment = number_alignment
            cell.border = thin_border

            cell = ws.cell(row=current_row, column=4, value=item['balance'])
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

            cell = ws.cell(row=current_row, column=5, value=item['reserved'])
            cell.alignment = number_alignment
            cell.border = thin_border
            if item['reserved'] > 0:
                cell.font = Font(color="F57C00", bold=True, size=10)
                cell.fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
            else:
                cell.font = data_font

            cell = ws.cell(row=current_row, column=6, value=item['available'])
            cell.alignment = number_alignment
            cell.border = thin_border
            cell.font = Font(color="1B5E20", bold=True, size=10)

            cell = ws.cell(row=current_row, column=7, value=None)
            cell.alignment = number_alignment
            cell.border = thin_border

            current_row += 1

        ws.row_dimensions[current_row].height = 8
        current_row += 1

    ws.freeze_panes = ws.cell(row=HEADER_ROW + 1, column=1)

    last_row = current_row + 1

    total_items = len(data)
    low_items = len([item for item in data if 0 < item['balance'] < 5])
    zero_items = len([item for item in data if item['balance'] <= 0])
    reserved_items = len([item for item in data if item['reserved'] > 0])

    info_font = Font(italic=True, size=9, color="777777")

    ws.cell(row=last_row, column=2, value=f'Дата выгрузки: {get_perm_time().strftime("%d.%m.%Y %H:%M")} (Пермь)').font = info_font
    ws.cell(row=last_row + 1, column=2, value=f'Всего позиций: {total_items}').font = info_font
    ws.cell(row=last_row + 2, column=2, value=f'Позиций с остатком менее 5 шт: {low_items}').font = info_font
    ws.cell(row=last_row + 3, column=2, value=f'Позиций с нулевым остатком: {zero_items}').font = info_font
    ws.cell(row=last_row + 4, column=2, value=f'Позиций в резерве: {reserved_items}').font = info_font
    ws.cell(row=last_row + 5, column=2, value=f'Выгрузил: {username}').font = info_font

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
    print("=" * 60)
    print("🚀 Запуск приложения...")
    print(f"📦 База данных: {DATABASE_URL[:50]}...")
    print(f"📁 Папка цен: {PRICES_DIR}")
    print(f"🔑 SECRET_KEY: {'*' * 20} (скрыт)")
    print("=" * 60)
    app.run(debug=True, host='0.0.0.0', port=5000)