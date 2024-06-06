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
            <html>
                <head>            
                <style>
                body {
    	background: #DCDDDF url(https://cssdeck.com/uploads/media/items/7/7AF2Qzt.png);
    	color: #000;
    	font: 14px Arial;
    	margin: 0 auto;
    	padding: 0;
    	position: relative;
    }
    h1{ font-size:28px;}
    h2{ font-size:26px;}
    h3{ font-size:18px;}
    h4{ font-size:16px;}
    h5{ font-size:14px;}
    h6{ font-size:12px;}
    h1,h2,h3,h4,h5,h6{ color:#563D64;}
    small{ font-size:10px;}
    b, strong{ font-weight:bold;}
    a{ text-decoration: none; }
    a:hover{ text-decoration: underline; }
    .left { float:left; }
    .right { float:right; }
    .alignleft { float: left; margin-right: 15px; }
    .alignright { float: right; margin-left: 15px; }
    .clearfix:after,
    form:after {
    	content: ".";
    	display: block;
    	height: 0;
    	clear: both;
    	visibility: hidden;
    }
    .container { margin: 25px auto; position: relative; width: 900px; }
    #content {
    	background: #f9f9f9;
    	background: -moz-linear-gradient(top,  rgba(248,248,248,1) 0%, rgba(249,249,249,1) 100%);
    	background: -webkit-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: -o-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: -ms-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#f8f8f8', endColorstr='#f9f9f9',GradientType=0 );
    	-webkit-box-shadow: 0 1px 0 #fff inset;
    	-moz-box-shadow: 0 1px 0 #fff inset;
    	-ms-box-shadow: 0 1px 0 #fff inset;
    	-o-box-shadow: 0 1px 0 #fff inset;
    	box-shadow: 0 1px 0 #fff inset;
    	border: 1px solid #c4c6ca;
    	margin: 0 auto;
    	padding: 25px 0 0;
    	position: relative;
    	text-align: center;
    	text-shadow: 0 1px 0 #fff;
    	width: 400px;
    }
    #content h1 {
    	color: #7E7E7E;
    	font: bold 25px Helvetica, Arial, sans-serif;
    	letter-spacing: -0.05em;
    	line-height: 20px;
    	margin: 10px 0 30px;
    }
    #content h1:before,
    #content h1:after {
    	content: "";
    	height: 1px;
    	position: absolute;
    	top: 10px;
    	width: 27%;
    }
    #content h1:after {
    	background: rgb(126,126,126);
    	background: -moz-linear-gradient(left,  rgba(126,126,126,1) 0%, rgba(255,255,255,1) 100%);
    	background: -webkit-linear-gradient(left,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: -o-linear-gradient(left,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: -ms-linear-gradient(left,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: linear-gradient(left,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
        right: 0;
    }
    #content h1:before {
    	background: rgb(126,126,126);
    	background: -moz-linear-gradient(right,  rgba(126,126,126,1) 0%, rgba(255,255,255,1) 100%);
    	background: -webkit-linear-gradient(right,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: -o-linear-gradient(right,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: -ms-linear-gradient(right,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
    	background: linear-gradient(right,  rgba(126,126,126,1) 0%,rgba(255,255,255,1) 100%);
        left: 0;
    }
    #content:after,
    #content:before {
    	background: #f9f9f9;
    	background: -moz-linear-gradient(top,  rgba(248,248,248,1) 0%, rgba(249,249,249,1) 100%);
    	background: -webkit-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: -o-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: -ms-linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	background: linear-gradient(top,  rgba(248,248,248,1) 0%,rgba(249,249,249,1) 100%);
    	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#f8f8f8', endColorstr='#f9f9f9',GradientType=0 );
    	border: 1px solid #c4c6ca;
    	content: "";
    	display: block;
    	height: 100%;
    	left: -1px;
    	position: absolute;
    	width: 100%;
    }
    #content:after {
    	-webkit-transform: rotate(2deg);
    	-moz-transform: rotate(2deg);
    	-ms-transform: rotate(2deg);
    	-o-transform: rotate(2deg);
    	transform: rotate(2deg);
    	top: 0;
    	z-index: -1;
    }
    #content:before {
    	-webkit-transform: rotate(-3deg);
    	-moz-transform: rotate(-3deg);
    	-ms-transform: rotate(-3deg);
    	-o-transform: rotate(-3deg);
    	transform: rotate(-3deg);
    	top: 0;
    	z-index: -2;
    }
    #content form { margin: 0 20px; position: relative }
    #content form input[type="text"],
    #content form input[type="password"] {
    	-webkit-border-radius: 3px;
    	-moz-border-radius: 3px;
    	-ms-border-radius: 3px;
    	-o-border-radius: 3px;
    	border-radius: 3px;
    	-webkit-box-shadow: 0 1px 0 #fff, 0 -2px 5px rgba(0,0,0,0.08) inset;
    	-moz-box-shadow: 0 1px 0 #fff, 0 -2px 5px rgba(0,0,0,0.08) inset;
    	-ms-box-shadow: 0 1px 0 #fff, 0 -2px 5px rgba(0,0,0,0.08) inset;
    	-o-box-shadow: 0 1px 0 #fff, 0 -2px 5px rgba(0,0,0,0.08) inset;
    	box-shadow: 0 1px 0 #fff, 0 -2px 5px rgba(0,0,0,0.08) inset;
    	-webkit-transition: all 0.5s ease;
    	-moz-transition: all 0.5s ease;
    	-ms-transition: all 0.5s ease;
    	-o-transition: all 0.5s ease;
    	transition: all 0.5s ease;
    	background: #eae7e7 url(https://cssdeck.com/uploads/media/items/8/8bcLQqF.png) no-repeat;
    	border: 1px solid #c8c8c8;
    	color: #777;
    	font: 13px Helvetica, Arial, sans-serif;
    	margin: 0 0 10px;
    	padding: 15px 10px 15px 40px;
    	width: 80%;
    }
    #content form input[type="text"]:focus,
    #content form input[type="password"]:focus {
    	-webkit-box-shadow: 0 0 2px #ed1c24 inset;
    	-moz-box-shadow: 0 0 2px #ed1c24 inset;
    	-ms-box-shadow: 0 0 2px #ed1c24 inset;
    	-o-box-shadow: 0 0 2px #ed1c24 inset;
    	box-shadow: 0 0 2px #ed1c24 inset;
    	background-color: #fff;
    	border: 1px solid #ed1c24;
    	outline: none;
    }
    #username { background-position: 10px 10px !important }
    #password { background-position: 10px -53px !important }
    #content form input[type="submit"] {
    	background: rgb(254,231,154);
    	background: -moz-linear-gradient(top,  rgba(254,231,154,1) 0%, rgba(254,193,81,1) 100%);
    	background: -webkit-linear-gradient(top,  rgba(254,231,154,1) 0%,rgba(254,193,81,1) 100%);
    	background: -o-linear-gradient(top,  rgba(254,231,154,1) 0%,rgba(254,193,81,1) 100%);
    	background: -ms-linear-gradient(top,  rgba(254,231,154,1) 0%,rgba(254,193,81,1) 100%);
    	background: linear-gradient(top,  rgba(254,231,154,1) 0%,rgba(254,193,81,1) 100%);
    	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#fee79a', endColorstr='#fec151',GradientType=0 );
    	-webkit-border-radius: 30px;
    	-moz-border-radius: 30px;
    	-ms-border-radius: 30px;
    	-o-border-radius: 30px;
    	border-radius: 30px;
    	-webkit-box-shadow: 0 1px 0 rgba(255,255,255,0.8) inset;
    	-moz-box-shadow: 0 1px 0 rgba(255,255,255,0.8) inset;
    	-ms-box-shadow: 0 1px 0 rgba(255,255,255,0.8) inset;
    	-o-box-shadow: 0 1px 0 rgba(255,255,255,0.8) inset;
    	box-shadow: 0 1px 0 rgba(255,255,255,0.8) inset;
    	border: 1px solid #D69E31;
    	color: #85592e;
    	cursor: pointer;
    	float: left;
    	font: bold 15px Helvetica, Arial, sans-serif;
    	height: 35px;
    	margin: 20px 0 35px 15px;
    	position: relative;
    	text-shadow: 0 1px 0 rgba(255,255,255,0.5);
    	width: 120px;
    }
    #content form input[type="submit"]:hover {
    	background: rgb(254,193,81);
    	background: -moz-linear-gradient(top,  rgba(254,193,81,1) 0%, rgba(254,231,154,1) 100%);
    	background: -webkit-linear-gradient(top,  rgba(254,193,81,1) 0%,rgba(254,231,154,1) 100%);
    	background: -o-linear-gradient(top,  rgba(254,193,81,1) 0%,rgba(254,231,154,1) 100%);
    	background: -ms-linear-gradient(top,  rgba(254,193,81,1) 0%,rgba(254,231,154,1) 100%);
    	background: linear-gradient(top,  rgba(254,193,81,1) 0%,rgba(254,231,154,1) 100%);
    	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#fec151', endColorstr='#fee79a',GradientType=0 );
    }
    #content form div a {
    	color: #004a80;
        float: right;
        font-size: 12px;
        margin: 30px 15px 0 0;
        text-decoration: underline;
    }
    .button {
    	background: rgb(247,249,250);
    	background: -moz-linear-gradient(top,  rgba(247,249,250,1) 0%, rgba(240,240,240,1) 100%);
    	background: -webkit-linear-gradient(top,  rgba(247,249,250,1) 0%,rgba(240,240,240,1) 100%);
    	background: -o-linear-gradient(top,  rgba(247,249,250,1) 0%,rgba(240,240,240,1) 100%);
    	background: -ms-linear-gradient(top,  rgba(247,249,250,1) 0%,rgba(240,240,240,1) 100%);
    	background: linear-gradient(top,  rgba(247,249,250,1) 0%,rgba(240,240,240,1) 100%);
    	filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#f7f9fa', endColorstr='#f0f0f0',GradientType=0 );
    	-webkit-box-shadow: 0 1px 2px rgba(0,0,0,0.1) inset;
    	-moz-box-shadow: 0 1px 2px rgba(0,0,0,0.1) inset;
    	-ms-box-shadow: 0 1px 2px rgba(0,0,0,0.1) inset;
    	-o-box-shadow: 0 1px 2px rgba(0,0,0,0.1) inset;
    	box-shadow: 0 1px 2px rgba(0,0,0,0.1) inset;
    	-webkit-border-radius: 0 0 5px 5px;
    	-moz-border-radius: 0 0 5px 5px;
    	-o-border-radius: 0 0 5px 5px;
    	-ms-border-radius: 0 0 5px 5px;
    	border-radius: 0 0 5px 5px;
    	border-top: 1px solid #CFD5D9;
    	padding: 15px 0;
    }
    .button a {
    	background: url(https://cssdeck.com/uploads/media/items/8/8bcLQqF.png) 0 -112px no-repeat;
    	color: #7E7E7E;
    	font-size: 17px;
    	padding: 2px 0 2px 40px;
    	text-decoration: none;
    	-webkit-transition: all 0.3s ease;
    	-moz-transition: all 0.3s ease;
    	-ms-transition: all 0.3s ease;
    	-o-transition: all 0.3s ease;
    	transition: all 0.3s ease;
    }
    .button a:hover {
    	background-position: 0 -135px;
    	color: #00aeef;
    }
                </style>
                </head>
                <body>
                <div class="container">
    	<section id="content">
    		<form method=post enctype=multipart/form-data>
    		<br/>
    		<br/>
    			<h1>Загрузите файл Excel file</h1>
    			<div>
    				<input type="file" name="file">
    			</div>
    			<div>
    				<input type="submit" value="Построить">
    			</div>
    		</form><!-- form -->

    	</section><!-- content -->
    </div><!-- container -->
                </body>
            </html>
        '''


if __name__ == '__main__':
    app.run(debug=True)


