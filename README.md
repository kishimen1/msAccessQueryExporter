# ⚡ Access クエリエクスポーター

Access データベース (.accdb / .mdb) のテーブル構造をブラウザ上で可視化するウェブアプリです。

## 🌐 オンライン版（GitHub Pages）

**👉 [https://kishimen1.github.io/msAccessQueryExporter/](https://kishimen1.github.io/msAccessQueryExporter/)**

- ブラウザ上で完結（サーバー不要）
- Access ファイルをドラッグ＆ドロップするだけ
- テーブル一覧・フィールド情報を即座に表示

## 🖥️ ローカル版（Python / Windows専用）

ローカル版では DAO COM オートメーションを使用し、**クエリの SQL 定義**も含めた完全なエクスポートが可能です。

### 必要環境
- Windows
- Python 3.8+
- Microsoft Access Database Engine (ACE)

### セットアップ
```bash
pip install -r requirements.txt
python app.py
```

ブラウザで http://localhost:5000 を開いてご利用ください。

## 機能比較

| 機能 | オンライン版 | ローカル版 |
|------|:-----------:|:----------:|
| テーブル一覧 | ✅ | ✅ |
| フィールド情報 | ✅ | ✅ |
| リレーションシップ | ❌ | ✅ |
| クエリ SQL 出力 | ❌ | ✅ |
| 一括ダウンロード | ✅ | ✅ |
| サーバー不要 | ✅ | ❌ |

## ライセンス

MIT
