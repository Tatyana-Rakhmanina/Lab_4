import docx
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import glob
import sqlite3
import os

from PIL.Image import Image
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
birth TEXT NOT NULL,
sex TEXT NOT NULL,
email TEXT NOT NULL) ''')
cursor.execute('INSERT INTO Analyst  VALUES (?,?,?,?,?)', (1, "Иванов Иван Иванович","1999.09.08","мужской","ffff@gmail.com"))
connection.commit()

# Создаем таблицу Types_of_visualization
cursor.execute(''' CREATE TABLE IF NOT EXISTS Types_of_visualization(
id_ types_of_visualization INTEGER PRIMARY KEY,
name TEXT NOT NULL) ''')

# Создаем таблицу Data
cursor.execute(''' CREATE TABLE IF NOT EXISTS Data (
id_data INTEGER PRIMARY KEY,
segment TEXT,
trade_group TEXT,
date TEXT NOT NULL,
sales REAL,
sales_rrp REAL,
markdown REAL,
markdown_perc REAL,
markup REAL,
inv REAL) ''')

# Создаем таблицу Report
cursor.execute(''' CREATE TABLE IF NOT EXISTS Report (
id_report INTEGER PRIMARY KEY,
id_data INTEGER,
id_analyst INTEGER ,
FOREIGN KEY (id_data) REFERENCES Date (id_data),
FOREIGN KEY (id_analyst) REFERENCES Analyst (id_analyst)) ''')

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

# Сохранение графиков в файл Word
@app.route('/save-report')
def save_report(images):
    output_dir = "output_folder"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    for i, image_path in enumerate(images):
        img = Image.open(image_path)
        doc = docx.Document()
        doc.add_picture(image_path)
        doc.save(f"{output_dir}/image_{i + 1}.docx")

    return 'Отчет успешно сохранен в файле report.docx'


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
            cursor.execute('INSERT INTO Data (segment,trade_group,date,sales,markdown,markdown_perc,sales_rrp,markup,inv) VALUES (?,?,?,?,?,?,?,?,?)', (data['Дата'], data['Сегмент'], data['Товарная группа'], data['Продажа в оц'],data['Цена в РРЦ'], data['Скидка'],data['Скидка%'], data['Наценка'],data['Остаток руб']))
            connection.commit()
            # График
            ax=plt.plot(data['Дата'], data['Продажа в оц'])
            plt.xlabel('Дата')
            plt.ylabel('Продажа в оц')
            plt.title('График продаж в динамике')
            plt.savefig('D:/result/plot1.jpg', format='jpg')
            plt.show()
            # График
            ax = plt.plot(data['Дата'], data['Наценка'])
            plt.xlabel('Дата')
            plt.ylabel('Наценка')
            plt.title('График продаж в динамике')
            plt.savefig('D:/result/plot2.jpg', format='jpg')
            plt.show()
            # График
            ax = plt.plot(data['Дата'], data['Остатки руб'])
            plt.xlabel('Дата')
            plt.ylabel('Остатки руб')
            plt.title('График продаж в динамике')
            plt.savefig('D:/result/plot3.jpg', format='jpg')
            plt.show()
            # Столбиковая диаграмма
            plt.bar(data['Дата'], data['Цена в РРЦ'])
            plt.xlabel('Дата')
            plt.ylabel('Цена в РРЦ')
            plt.title('Столбиковая диаграмма цен в РРЦ')
            plt.savefig('D:/result/plot4.jpg', format='jpg')
            plt.show()
            # Столбиковая диаграмма
            plt.bar(data['Дата'], data['Скидка'])
            plt.xlabel('Дата')
            plt.ylabel('Скидка')
            plt.title('Столбиковая диаграмма цен в РРЦ')
            plt.savefig('D:/result/plot5.jpg', format='jpg')
            plt.show()
            # Столбиковая диаграмма
            plt.bar(data['Дата'], data['Скидка%'])
            plt.xlabel('Дата')
            plt.ylabel('Скидка%')
            plt.title('Столбиковая диаграмма цен в РРЦ')
            plt.savefig('D:/result/plot6.jpg', format='jpg')
            plt.show()
            # Рассеивание
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.scatter(x=data['Остаток руб'], y=data['Цена в РРЦ'])
            plt.xlabel("Остаток руб")
            plt.ylabel("Цена в РРЦ")
            plt.savefig('D:/result/plot7.jpg', format='jpg')
            plt.show()
            # Круговая диаграмма
            fig1, ax1 = plt.subplots()
            ax1.pie(data['Скидка'], labels=data['Дата'], autopct='%1.1f%%')
            plt.savefig('D:/result/plot8.jpg', format='jpg')
            plt.show()
            # Круговая диаграмма
            fig1, ax1 = plt.subplots()
            ax1.pie(data['Скидка%'], labels=data['Дата'], autopct='%1.1f%%')
            plt.savefig('D:/result/plot9.jpg', format='jpg')
            plt.show()
            # Круговая диаграмма
            fig1, ax1 = plt.subplots()
            ax1.pie(data['Продажа в оц'], labels=data['Дата'], autopct='%1.1f%%')
            plt.savefig('D:/result/plot10.jpg', format='jpg')
            plt.show()
            images = ["D:/result/plot1.jpg", "D:/result/plot2.jpg", "D:/result/plot3.jpg","D:/result/plot4.jpg", "D:/result/plot5.jpg", "D:/result/plot6.jpg", "D:/result/plot7.jpg", "D:/result/plot8.jpg","D:/result/plot9.jpg", "D:/result/plot10.jpg"]
            cursor.execute('SELECT id_data FROM Data ORDER BY id DESC LIMIT 1')
            id_data = cursor.fetchone()
            cursor.execute('INSERT INTO Report (id_data,id_analyst) VALUES (?,?)',
                           (id_data, 1))
            connection.commit()
            save_report(images)

            return redirect(url_for('static'))

    return '''
 <!doctype html>
<title>Загрузка файла</title>
<h1>Загрузите файл Excel</h1>
<form method=post enctype=multipart/form-data>
<br/>
  <input type=file name=file>
  <input type=submit value=Загрузить файл>
</form>
'''


if __name__ == '__main__':
    app.run(debug=True)


