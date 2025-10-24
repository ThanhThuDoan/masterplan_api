from flask import Flask, jsonify, request
import os
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from werkzeug.utils import secure_filename

app = Flask(__name__)

# -------------------- CẤU HÌNH --------------------
UPLOAD_FOLDER = "/tmp"  # Render chỉ cho phép ghi tạm trong /tmp
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# File credentials.json nên để trong Render Environment variable
CREDENTIALS_ENV = "GOOGLE_CREDS"  # tên biến môi trường

def get_credentials_from_env():
    import json
    creds_json = os.environ.get(CREDENTIALS_ENV)
    if not creds_json:
        raise Exception("Thiếu biến môi trường GOOGLE_CREDS trên Render!")
    creds_dict = json.loads(creds_json)
    with open("/tmp/credentials.json", "w") as f:
        json.dump(creds_dict, f)
    return "/tmp/credentials.json"

# -------------------- HÀM PHỤ --------------------
def connect_to_google_sheets(credentials_file, scope):
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)
    return client

def update_google_sheet(worksheet, dataframe):
    dataframe = dataframe.where(pd.notnull(dataframe), "")
    header = [dataframe.columns.values.tolist()]
    data = dataframe.values.tolist()
    worksheet.batch_clear(['A4:ZZ'])
    worksheet.update('A4', header)
    worksheet.update('A5', data)
    return len(data)

# -------------------- API CHÍNH --------------------
@app.route('/run_upload', methods=['POST'])
def run_upload():
    try:
        if 'file' not in request.files:
            return jsonify({"status": "error", "message": "Không có file Excel gửi kèm!"}), 400

        file = request.files['file']
        filename = secure_filename(file.filename)
        save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(save_path)

        # Kết nối Google Sheets
        credentials_file = get_credentials_from_env()
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/11qBWD7ew70L-MTVY0la5Fo4JxfmGbir_DR3vzB28u_8/edit#gid=826240624"
        worksheet_name = "DATABASE"

        client = connect_to_google_sheets(credentials_file, scope)
        worksheet = client.open_by_url(spreadsheet_url).worksheet(worksheet_name)

        # Đọc file Excel
        df = pd.read_excel(save_path, sheet_name="Master Plan", header=2)

        # Ghi dữ liệu lên Google Sheets
        rows = update_google_sheet(worksheet, df)

        return jsonify({"status": "success", "message": f"Đã upload {rows} dòng từ {filename} lên Google Sheets."})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/')
def home():
    return jsonify({"message": "Masterplan API hoạt động bình thường 🚀"})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
