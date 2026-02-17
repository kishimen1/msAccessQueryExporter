# ⚡ Access クエリエクスポーター

Access データベース (.accdb / .mdb) のテーブル構造・クエリ SQL をブラウザ上で可視化するツールです。

## 🌐 オンライン版（GitHub Pages）

**👉 [https://kishimen1.github.io/msAccessQueryExporter/](https://kishimen1.github.io/msAccessQueryExporter/)**

- ブラウザ上で完結（サーバー不要）
- Access ファイルをドラッグ＆ドロップするだけ
- テーブル一覧・フィールド情報を即座に表示

## 🖥️ ローカル版（Python不要 / Windows専用）

`local/` フォルダに格納されています。**Python のインストールは不要**で、ドラッグ＆ドロップで簡単に使えます。

### 必要環境
- Windows
- Microsoft Access または Access Database Engine (ACE)

### 使い方

1. `local/` フォルダごと PC にコピー
2. **Access ファイル (.accdb / .mdb) を `AccessQueryExporter.bat` にドラッグ＆ドロップ**
3. 解析結果がブラウザで自動表示 & テキストファイルとして保存

> ダブルクリックで起動した場合は、ファイル選択ダイアログが表示されます。

### 出力ファイル
- `○○_result.json` — ビューア用データ（viewer.html で閲覧）
- `○○_クエリ一覧.txt` — テキスト形式のエクスポート

---

<details>
<summary>📦 Python版ローカルサーバー（旧版）</summary>

Python + Flask を使ったWebサーバー版です。

#### 必要環境
- Windows / Python 3.8+ / Microsoft Access Database Engine (ACE)

#### セットアップ
```bash
pip install -r requirements.txt
python app.py
```

ブラウザで http://localhost:5000 を開いてご利用ください。

</details>

## 機能比較

| 機能 | オンライン版 | ローカル版 | Python版（旧） |
|------|:-----------:|:----------:|:--------------:|
| テーブル一覧 | ✅ | ✅ | ✅ |
| フィールド情報 | ✅ | ✅ | ✅ |
| リレーションシップ | ❌ | ✅ | ✅ |
| クエリ SQL 出力 | ❌ | ✅ | ✅ |
| 一括ダウンロード | ✅ | ✅ | ✅ |
| Python 不要 | ✅ | ✅ | ❌ |
| サーバー不要 | ✅ | ✅ | ❌ |

## ライセンス

MIT
