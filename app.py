from flask import Flask, request, render_template, send_file , jsonify

import pandas as pd

import os



app = Flask(__name__)





EXCEL_FILE = 'data.xlsx'

Personel_FILE = 'اطلاعات پرسنل.xlsx'



@app.route('/')

def index():

    return render_template('index.html')





@app.route('/add', methods=['POST'])

def add_entry():

    value = request.json.get('value')

    # int checking

    try:

        normalized_value = int(value)

    except ValueError:

        return jsonify({"message": "فرمت کد ملی صحیح نمیباشد!"}), 400

    

    if os.path.exists(Personel_FILE):

        pf = pd.read_excel(Personel_FILE)

    else:

        return jsonify({"message": "فایل اطلاعات پرسنل موجود نمیباشد!"}), 400

    

    if os.path.exists(EXCEL_FILE):

        df = pd.read_excel(EXCEL_FILE)

    else:

        df = pd.DataFrame(columns=['ID'])





    pf['شماره ملی '] = pd.to_numeric(pf['شماره ملی '])

    pf.dropna(subset=['شماره ملی '], inplace=True)

    # print(pf['شماره ملی '][5])



    if normalized_value in pf['شماره ملی '].values:

        df['ID'] = pd.to_numeric(df['ID'], errors='coerce')

        df.dropna(subset=['ID'], inplace=True)

        if normalized_value not in df['ID'].values:

            new_row = pd.DataFrame({'ID': [normalized_value]})

            df = pd.concat([df, new_row], ignore_index=True)

            df.to_excel(EXCEL_FILE, index=False)

            return jsonify({"message": "ثبت شد"}), 201

        else:

            return jsonify({"message": "این کد ملی قبلا ثبت شده"}), 409 



    else:

        return jsonify({"message": "این کد ملی در اطلاعات پرسنلی موجود نیست!"}), 409





@app.route('/export', methods=['GET'])

def export_to_excel():

    if os.path.exists(EXCEL_FILE):

        return send_file(EXCEL_FILE, as_attachment=True)

    return "فایلی جهت دانلود وجود ندارد!", 404



if __name__ == '__main__':

    app.run(debug=True)
