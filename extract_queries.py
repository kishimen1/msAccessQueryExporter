# -*- coding: utf-8 -*-
"""
Access データベース クエリ一括エクスポート (Python版)

使用方法:
    python extract_queries.py "C:\path\to\database.accdb"

必要なライブラリ:
    pip install pywin32

注意:
    - Microsoft Access Database Engine (ACE) が必要
    - 32bit Python には 32bit ACE、64bit Python には 64bit ACE が必要
"""

import sys
import os
from datetime import datetime

def extract_queries(db_path: str) -> None:
    """Accessデータベースからクエリ定義を抽出してテキストファイルに出力"""

    try:
        import win32com.client
    except ImportError:
        print("エラー: pywin32 がインストールされていません。")
        print("インストール: pip install pywin32")
        sys.exit(1)

    # 出力ファイルパスを生成
    base_name = os.path.splitext(db_path)[0]
    output_path = f"{base_name}_クエリ一覧.txt"

    try:
        # DAO経由でデータベースを開く
        engine = win32com.client.Dispatch("DAO.DBEngine.120")
        db = engine.OpenDatabase(db_path)
    except Exception as e:
        print(f"エラー: データベースを開けませんでした。")
        print(f"パス: {db_path}")
        print(f"詳細: {e}")
        sys.exit(1)

    lines = []

    # ヘッダー
    lines.append("=" * 50)
    lines.append(f"データベース: {os.path.basename(db_path)}")
    lines.append(f"出力日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 50)
    lines.append("")

    # テーブル一覧
    lines.append("■■■ テーブル一覧 ■■■")
    lines.append("")
    for tdef in db.TableDefs:
        if not tdef.Name.startswith("MSys") and not tdef.Name.startswith("~"):
            lines.append(f"  ・{tdef.Name}")
    lines.append("")

    # リレーションシップ
    lines.append("■■■ リレーションシップ ■■■")
    lines.append("")
    for rel in db.Relations:
        lines.append(f"  {rel.Table} → {rel.ForeignTable}")
    lines.append("")

    # クエリ一覧とSQL
    lines.append("■■■ クエリ一覧とSQL定義 ■■■")
    lines.append("")

    query_count = 0
    for qdef in db.QueryDefs:
        if not qdef.Name.startswith("~"):
            query_count += 1
            lines.append(f"【{query_count}】 {qdef.Name}")
            lines.append("-" * 50)
            lines.append(qdef.SQL)
            lines.append("")
            lines.append("")

    lines.append("=" * 50)
    lines.append(f"総クエリ数: {query_count}")
    lines.append("=" * 50)

    db.Close()

    # ファイル出力 (UTF-8)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("出力完了！")
    print(f"クエリ数: {query_count}")
    print(f"出力先: {output_path}")


def main():
    if len(sys.argv) < 2:
        print("使用方法: python extract_queries.py \"C:\\path\\to\\database.accdb\"")
        sys.exit(1)

    db_path = sys.argv[1]

    if not os.path.exists(db_path):
        print(f"エラー: ファイルが見つかりません: {db_path}")
        sys.exit(1)

    extract_queries(db_path)


if __name__ == "__main__":
    main()
