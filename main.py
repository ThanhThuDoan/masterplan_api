from flask import Flask, jsonify, request
import sys
import os
import glob
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ================== KHỞI TẠO APP FLASK ==================
app = Flask(__name__)

# ================== CÁC HÀM GỐC (giữ nguyên code của bạn) ==================
def connect_to_google_sheets(credentials_file, scope):
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)
    return client

def get_worksheet(client, spreadsheet_url, worksheet_name):
    spreadsheet = client.open_by_url(spreadsheet_url)
    worksheet = spreadsheet.worksheet(worksheet_name)
    return worksheet

def read_excel_data(file_path, sheet_name="Master Plan", selected_columns=None):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=2)
    for col in df.columns:
        if "CHD" in col:
            df.rename(columns={col: "CHD"}, inplace=True)
            break
    available_cols = [col for col in selected_columns.keys() if col in df.columns]
    df = df[available_cols]
    df = df.rename({col: selected_columns[col] for col in available_cols}, axis=1)
    return df

def update_google_sheet(worksheet, dataframe):
    dataframe = dataframe.where(pd.notnull(dataframe), "")
    header = [dataframe.columns.values.tolist()]
    data = dataframe.values.tolist()
    worksheet.batch_clear(['A4:ZZ'])
    worksheet.update('A4', header)
    worksheet.update('A5', data)
    return len(data)

# ================== API CHÍNH ==================
@app.route('/run_upload', methods=['POST'])
def run_upload():
    try:
        credentials_file = "credentials.json"   # file này bạn upload cùng code lên Render
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        spreadsheet_url = "https://docs.google.com/spreadsheets/d/11qBWD7ew70L-MTVY0la5Fo4JxfmGbir_DR3vzB28u_8/edit?gid=826240624#gid=826240624"
        worksheet_name = 'DATABASE'
        folder_path = r"\\rg2fnp01.os.crystal.com\ShareRoot\SharedFolder\REGENT\DECATHLON\PUBLIC\2. DKL OM\1. KPI Management\5. Master plan"
        excel_sheet_name = "Master Plan"

        selected_columns = {
            "CC": "Style",
            "R3": "Model",
            "Season": "Season",
            "Buy": "Buy Date",
            "ORDER NO": "PO",
            "CHD": "CHD",
            "EHD": "EHD",
            "HOT": "HOT",
            "Standard Delay Reason": "Delay Reason",
            "CAC": "CAC Code",
            "Firm order": "Order Qty"
        }

        client = connect_to_google_sheets(credentials_file, scope)
        worksheet = get_worksheet(client, spreadsheet_url, worksheet_name)

        # Tìm file Excel
        file_patterns = ["AW24*", "AW25*", "SS25*", "SS26*"]
        excel_files = []
        for pattern in file_patterns:
            excel_files.extend(glob.glob(os.path.join(folder_path, f"{pattern}.xlsx")))

        if not excel_files:
            return jsonify({"status": "error", "message": "Không tìm thấy file Excel phù hợp."}), 404

        all_dfs = [read_excel_data(f, excel_sheet_name, selected_columns) for f in excel_files]
        final_df = pd.concat(all_dfs, ignore_index=True)
        rows = update_google_sheet(worksheet, final_df)

        return jsonify({"status": "success", "message": f"Đã upload {rows} dòng lên Google Sheets."})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/')
def home():
    return jsonify({"message": "Masterplan API đang hoạt động!"})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)