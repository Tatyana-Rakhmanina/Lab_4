# Импортируем необходимые модули
'''

from flask import Flask, render_template
import matplotlib.pyplot as plt
import io
import base64

app = Flask(__name__)


@app.route('/')
def index():
    # Генерируем данные для графика (здесь просто пример)
    x = [1, 2, 3, 4, 5]
    y = [10, 15, 7, 10, 13]

    # Создаем график
    plt.plot(x, y)

    # Сохраняем график в байтовом объекте
    img = io.BytesIO()
    plt.savefig(img, format='png')
    img.seek(0)

    # Кодируем байтовый объект в base64
    plot_url = base64.b64encode(img.getvalue()).decode()

    # Отрисовываем шаблон страницы с графиком
    return render_template('index.html', plot_url=plot_url)


if __name__ == '__main__':
    app.run(debug=True)

    return
   <!doctype html >
    < title > Загрузка файла < / title >
    < h1 > Загрузите файл Excel </h1 >
    < form
    method = post
    enctype = multipart / form - data >
    < input
    type = file
    name = file >
    < input
    type = submit
    value = Построить >
< / form >

    html
    <!DOCTYPE html >
< html lang = "en" >
< head >
< meta
charset = "UTF-8" >
< title > Graph < / title >
< / head >
< body >
< h1 > График < / h1 >

< !-- Отображение
графика
на
странице -->
< img
src = "data:image/png;base64,{{ plot_url }}" alt ="График">

</body>
</html>
'''
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import glob
import sqlite3
import os
from flask import Flask, request, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Указываем путь для загрузки файлов
app.config['UPLOAD_FOLDER'] = 'D:/upload/'
allowed_extensions = {'xlsx', 'xls'}

# Устанавливаем соединение с базой данных
connection = sqlite3.connect('my_database.db')
cursor = connection.cursor()

# Создаем таблицу Users
cursor.execute(''' CREATE TABLE IF NOT EXISTS Analyst (
id_analyst INTEGER PRIMARY KEY,
fio TEXT NOT NULL,
birth DATE NOT NULL,
sex INTEGER NOT NULL,
email TEXT NOT NULL) ''')



# Создаем таблицу Types_of_visualization
cursor.execute(''' CREATE TABLE IF NOT EXISTS Types_of_visualization(
id_ types_of_visualization INTEGER PRIMARY KEY,
name TEXT NOT NULL) ''')

# Создаем таблицу Data
cursor.execute(''' CREATE TABLE IF NOT EXISTS Data (
id_data INTEGER PRIMARY KEY,
measure TEXT,
date DATE NOT NULL,
sales REAL,
sales_rrp REAL,
markdown REAL,
markdown_perc REAL,
markup REAL,
inv REAL,
FOREIGN KEY (id_measure) REFERENCES Measure (id_measure)) ''')

# Создаем таблицу Report
cursor.execute(''' CREATE TABLE IF NOT EXISTS Report (
id_report INTEGER PRIMARY KEY,
id_measure INTEGER NOT NULL,
id_data INTEGER NOT NULL,
id_analyst INTEGER NOT NULL,
FOREIGN KEY (id_measure) REFERENCES Measure (id_measure),
FOREIGN KEY (id_data) REFERENCES Date (id_measure),
FOREIGN KEY (id_analyst) REFERENCES Analyst (id_measure)) ''')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def check_excel_format(file_path):
    try:
        df = pd.read_excel(file_path)
        expected_columns = ['Дата', 'Сегмент', 'Товарная группа', 'Продажа в оц','Цена в РРЦ', 'Скидка','Скидка%', 'Наценка','Остаток руб']
        if list(df.columns) == expected_columns:
            print("Файл имеет формат Excel и названия столбцов соответствуют ожидаемым.")
        else:
            print("Названия столбцов не соответствуют ожидаемым.")
    except Exception as e:
        print("Произошла ошибка при чтении файла:", e)



@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            check_excel_format(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            data = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # Сохранение графиков в базе данных

            cursor.execute('INSERT INTO Data (graph) VALUES (?)', (data['Дата'], data['Сегмент'], data['Товарная группа'], data['Продажа в оц'],data['Цена в РРЦ'], data['Скидка'],data['Скидка%'], data['Наценка'],data['Остаток руб']))
            connection.commit()

            # График
            ax=plt.plot(data['Дата'], data['Продажа в оц'])
            plt.xlabel('Дата')
            plt.ylabel('Продажа в оц')
            plt.title('График продаж в динамике')
            plt.show()
            # Столбиковая диаграмма
            plt.bar(data['Дата'], data['Цена в РРЦ'])
            plt.xlabel('Дата')
            plt.ylabel('Цена в РРЦ')
            plt.title('Столбиковая диаграмма цен в РРЦ')
            plt.show()

            # Рассеивание
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.scatter(x=data['Остаток руб'], y=data['Цена в РРЦ'])
            plt.xlabel("Остаток руб")
            plt.ylabel("Цена в РРЦ")

            plt.show()

            # Круговая диаграмма
            fig1, ax1 = plt.subplots()
            ax1.pie(data['Скидка'], labels=data['Дата'], autopct='%1.1f%%')
            plt.show()

            f= glob.glob('D:\static')
            ax.get_figure().savefig(os.path.splitext(f)[0] + '.png')
            plt.savefig('D:/static/plot.png')

            return redirect(url_for('static'))

    return '''
 <!doctype html>
<title font-size:25px;
	color:#D6CFCB>Загрузка файла</title>
<h1>Загрузите файл Excel</h1>
<form method=post enctype=multipart/form-data>
<br/>
  <input type=file name=file>
  <input type=submit value=Загрузить файл>
</form>
'''


if __name__ == '__main__':
    app.run(debug=True)


