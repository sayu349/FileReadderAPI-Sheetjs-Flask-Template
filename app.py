from flask import Flask
from flask import render_template
from flask import request

app = Flask(__name__)

# ============================================================
# ルーティング
# ============================================================
@app.route("/", methods=["POST","GET"])
def index_page():
    # POST
    if request.method == "POST":
        # xlsxファイル取得
        file = request.files["fileInput"]
        # シート名取得
        sheetNameSelectBox = request.form["sheetNameSelectBox"]
        # 行番号
        rowNumberInput = request.form["rowNumberInput"]
        # 金額列名
        columnNameSelectBox = request.form["columnNameSelectBox"]
        return render_template("download.html")
    # GET
    else:
        return render_template("index.html")


# ============================================================
# 実行
# ============================================================
if __name__== '__main__':
    app.run(debug=True)