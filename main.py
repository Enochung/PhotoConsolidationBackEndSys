from flask import Flask, request, jsonify
from flask_cors import CORS
from docx import Document
import os
import time
from flask import send_file

# 初始化 Flask 應用程式
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "http://localhost:4200"}})

# 設置上傳目錄
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/flaskapi/api/testtt')
def index():
    return jsonify({'name:':'enoch'},
                   {'email:': 'qqxxx@gmail.com'}), 200

# 定義處理圖片和說明上傳的 API 路由
@app.route('/flaskapi/api/upload', methods=['POST'])
def upload_file():
    try:
        # 檢查請求中是否包含圖片
        if 'images' not in request.files:
            return jsonify({"error": "未上傳圖片"}), 400

        # 獲取圖片列表和說明
        image_files = request.files.getlist('images')
        title = request.form.get('title', '').strip() or '預設標題'
        description = request.form.get('description', '').strip() or '無提供說明'
        shooting_time = request.form.get('shooting_time', '').strip() or f'{time.strftime("%Y%m%d")}'
        shooting_location = request.form.get('shooting_location', '').strip() or '台灣'
        photographer = request.form.get('photographer', '').strip() or 'None'

        # 創建 Word 文件並插入標題和說明
        document = Document()
        document.add_heading('圖片報告', 0)
        document.add_paragraph(f'標題: {title}')

        # 插入表格，每張圖片對應兩行，第一行放圖片，第二行放編號、說明等資訊
        for index, image_file in enumerate(image_files):
            image_path = os.path.join(UPLOAD_FOLDER, image_file.filename)
            image_file.save(image_path)

            # 創建一個 3 行 6 列的表格
            table = document.add_table(rows=3, cols=6)
            table.style = 'Table Grid'

            # 第一行，合併所有列來插入圖片
            row1 = table.rows[0]
            row1_cells = row1.cells
            cell1 = row1_cells[0].merge(row1_cells[5])  # 合併所有單元格
            paragraph = cell1.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(image_path, width=5000000)  # 插入圖片

            # 第二行，填寫表格資訊
            table.cell(1, 0).text = '照片編號'  # 照片編號
            table.cell(1, 1).text = f'{index + 1:02d}'
            table.cell(1, 2).text = '說明'  # 說明
            row2 = table.rows[1]
            cell2 = row2.cells[3].merge(row2.cells[5])
            cell2.text = f'{description}'

            # 第三行，攝影時間、地點和人員資訊
            table.cell(2, 0).text = '攝影時間'  # 攝影時間
            table.cell(2, 1).text = f'{shooting_time}'
            table.cell(2, 2).text = '攝影地點'  # 攝影地點
            table.cell(2, 3).text = f'{shooting_location}'
            table.cell(2, 4).text = '攝影人'  # 攝影人
            table.cell(2, 5).text = f'{photographer}'  # 攝影人

            # 每插入兩個表格後，添加一個段落作為標題
            if (index + 1) % 2 == 0 and (index + 1) != len(image_files):
                document.add_page_break()
                document.add_paragraph(f'標題: {title}')

        word_file_path = os.path.join(UPLOAD_FOLDER, f'{title}_{int(time.time())}.docx')

        document.save(word_file_path)

        delete_file_in_folder(UPLOAD_FOLDER)

        return jsonify({"message": "文件已成功處理", "word_file": word_file_path}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def delete_file_in_folder(folder_path, specific_file=None):
    try:
        if specific_file:
            file_path = os.path.join(folder_path, specific_file)
            if os.path.exists(file_path):  # 確保檔案存在
                os.remove(file_path)
                print(f"成功刪除指定檔案: {file_path}")
                return {"message": f"成功刪除指定檔案: {specific_file}"}
            else:
                return {"error": f"檔案 {specific_file} 不存在"}
        else:
            deleted_files = []
            for file_name in os.listdir(folder_path):
                file_path = os.path.join(folder_path, file_name)
                if file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):  # 圖片檔案類型
                    os.remove(file_path)
                    deleted_files.append(file_name)
                    print(f"刪除圖片檔案: {file_path}")
            if deleted_files:
                return {"message": "成功刪除所有圖片檔案", "deleted_files": deleted_files}
            else:
                return {"message": "沒有圖片檔案需要刪除"}
    except Exception as e:
        print(f"刪除圖片文件時出現錯誤: {str(e)}")

# 定義處理文件瀏覽的 API 路由
@app.route('/flaskapi/api/files', methods=['GET'])
def list_docx_files():
    try:
        docx_files = [f for f in os.listdir(UPLOAD_FOLDER) if f.endswith('.docx')]
        return jsonify({"docx_files": docx_files}), 200
    except FileNotFoundError:
        return jsonify({"error": "文件未找到"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 定義處理文件下載的 API 路由
@app.route('/flaskapi/api/download/<filename>', methods=['POST'])
def download_file(filename):
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if not os.path.exists(file_path):
            return jsonify({"error": "文件不存在"}), 404
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "文件未找到"}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# 定義處理文件刪除的 API 路由
@app.route('/flaskapi/api/delete/<filename>', methods=['POST'])
def delete_file(filename):
    try:
        result = delete_file_in_folder(UPLOAD_FOLDER, filename)
        if "error" in result:
            return jsonify(result), 404
        return jsonify(result), 200
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
