import os
from flask import Flask, render_template, jsonify, request, session, redirect, url_for, flash
import pandas as pd

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-me-please')  # поменяй в окружении

# Настройки
EXCEL_PATH = os.environ.get('EXCEL_PATH', 'data/report.xlsx')
EMPLOYEE_COLUMN = os.environ.get('EMPLOYEE_COLUMN', 'Испонитель')
PORT = int(os.environ.get('PORT', 5000))
APP_PASSWORD = os.environ.get('APP_PASSWORD', 'admin')  # обязательно поменяй

def load_stats():
    try:
        df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
    except Exception as e:
        return {'error': f'Не удалось открыть Excel: {e}', 'data': []}

    # Ищем колонку с сотрудниками
    col = None
    candidates = [EMPLOYEE_COLUMN, 'Сотрудник', 'ФИО', 'Name', 'Employee']
    for c in df.columns:
        if str(c).strip() in candidates:
            col = c
            break
    if col is None:
        for c in df.columns:
            low = str(c).lower()
            if 'сотр' in low or 'фио' in low or 'name' in low or 'employee' in low:
                col = c
                break
    if col is None:
        return {'error': 'Не найден столбец с именем сотрудника.', 'data': []}

    stats = df.groupby(col).size().reset_index(name='count').sort_values('count', ascending=False)
    stats = stats.rename(columns={col: 'employee'})
    records = stats.to_dict(orient='records')
    return {'error': None, 'data': records}

@app.route('/data')
def data():
    if not session.get('logged_in'):
        return jsonify({'error': 'auth', 'message': 'not logged in'}), 401
    res = load_stats()
    if res['error']:
        return jsonify({'error': res['error'], 'data': []})
    return jsonify({'error': None, 'data': res['data']})

@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return render_template('index.html')

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        pwd = request.form.get('password', '')
        if pwd == APP_PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            flash('Неверный пароль')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=PORT)

