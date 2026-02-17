# -*- coding: utf-8 -*-
"""
Access クエリ一括エクスポート Web アプリケーション

起動方法:
    python app.py

ブラウザで http://localhost:5000 を開いてご利用ください。
"""

import os
import sys
import uuid
import tempfile
import traceback
from datetime import datetime

from flask import Flask, request, jsonify, render_template

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 500 * 1024 * 1024  # 500MB上限

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "access_query_extractor")
os.makedirs(UPLOAD_DIR, exist_ok=True)


def get_query_type_label(qtype: int) -> str:
    """クエリタイプ番号を日本語ラベルに変換"""
    types = {
        0: "選択",
        1: "クロス集計",
        2: "削除",
        3: "更新",
        4: "追加",
        5: "テーブル作成",
        6: "データ定義",
        7: "パススルー",
        8: "ユニオン",
        9: "サブフォーム/サブレポート",
    }
    return types.get(qtype, f"その他({qtype})")


def extract_from_access(db_path: str) -> dict:
    """DAO を使って Access データベースからクエリ定義を抽出"""
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()

    try:
        # DAO エンジンを取得
        try:
            engine = win32com.client.Dispatch("DAO.DBEngine.120")
        except Exception:
            engine = win32com.client.Dispatch("DAO.DBEngine.36")

        db = engine.OpenDatabase(db_path)

        result = {
            "filename": os.path.basename(db_path),
            "exported_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "tables": [],
            "relationships": [],
            "queries": [],
        }

        # テーブル一覧
        for i in range(db.TableDefs.Count):
            tdef = db.TableDefs(i)
            name = tdef.Name
            if not name.startswith("MSys") and not name.startswith("~"):
                fields = []
                for j in range(tdef.Fields.Count):
                    field = tdef.Fields(j)
                    fields.append(field.Name)
                result["tables"].append({"name": name, "fields": fields})

        # リレーションシップ
        for i in range(db.Relations.Count):
            rel = db.Relations(i)
            fields_info = []
            for j in range(rel.Fields.Count):
                f = rel.Fields(j)
                fields_info.append(
                    {"from_field": f.Name, "to_field": f.ForeignName}
                )
            result["relationships"].append(
                {
                    "table": rel.Table,
                    "foreign_table": rel.ForeignTable,
                    "fields": fields_info,
                }
            )

        # クエリ一覧と SQL
        for i in range(db.QueryDefs.Count):
            qdef = db.QueryDefs(i)
            name = qdef.Name
            if not name.startswith("~"):
                result["queries"].append(
                    {
                        "name": name,
                        "sql": qdef.SQL.strip(),
                        "type": get_query_type_label(qdef.Type),
                    }
                )

        db.Close()
        return result

    finally:
        pythoncom.CoUninitialize()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "ファイルが選択されていません。"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "ファイルが選択されていません。"}), 400

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in (".accdb", ".mdb"):
        return jsonify({"error": "対応ファイル形式: .accdb, .mdb"}), 400

    # 一時ファイルに保存
    temp_name = f"{uuid.uuid4()}{ext}"
    temp_path = os.path.join(UPLOAD_DIR, temp_name)

    try:
        file.save(temp_path)
        result = extract_from_access(temp_path)
        return jsonify(result)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"解析エラー: {str(e)}"}), 500
    finally:
        # 一時ファイル削除
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


if __name__ == "__main__":
    print("=" * 50)
    print("  Access クエリエクスポーター")
    print("  http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, port=5000)
