' =============================================================
' Access クエリ一括エクスポート VBScript
' =============================================================
' 使用方法（コマンドプロンプトから実行）:
'   cscript ExtractQueries.vbs "C:\path\to\database.accdb"
'
' または、.accdbファイルをこのスクリプトにドラッグ＆ドロップ
' =============================================================
Option Explicit

Dim engine, db, qdef, tdef, rel
Dim fso, outFile
Dim dbPath, outputPath
Dim queryCount

' 引数チェック
If WScript.Arguments.Count = 0 Then
    WScript.Echo "使用方法: cscript ExtractQueries.vbs ""C:\path\to\database.accdb"""
    WScript.Quit 1
End If

dbPath = WScript.Arguments(0)

' 出力ファイルパスを生成
outputPath = Replace(dbPath, ".accdb", "_クエリ一覧.txt")
If InStr(dbPath, ".accdb") = 0 Then
    outputPath = Replace(dbPath, ".mdb", "_クエリ一覧.txt")
End If

On Error Resume Next

' DAO経由でデータベースを開く
Set engine = CreateObject("DAO.DBEngine.120")
If Err.Number <> 0 Then
    ' ACE 2010を試す
    Err.Clear
    Set engine = CreateObject("DAO.DBEngine.36")
    If Err.Number <> 0 Then
        WScript.Echo "エラー: DAO.DBEngineが見つかりません。"
        WScript.Echo "Microsoft Access Database Engine をインストールしてください。"
        WScript.Quit 1
    End If
End If

Set db = engine.OpenDatabase(dbPath)
If Err.Number <> 0 Then
    WScript.Echo "エラー: データベースを開けませんでした。"
    WScript.Echo "パス: " & dbPath
    WScript.Echo Err.Description
    WScript.Quit 1
End If

On Error GoTo 0

' ファイル出力準備
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.CreateTextFile(outputPath, True, True)  ' Unicode出力

' ヘッダー
outFile.WriteLine "=============================================="
outFile.WriteLine "データベース: " & fso.GetFileName(dbPath)
outFile.WriteLine "出力日時: " & Now()
outFile.WriteLine "=============================================="
outFile.WriteLine ""

' テーブル一覧
outFile.WriteLine "■■■ テーブル一覧 ■■■"
outFile.WriteLine ""
For Each tdef In db.TableDefs
    If Left(tdef.Name, 4) <> "MSys" And Left(tdef.Name, 1) <> "~" Then
        outFile.WriteLine "  ・" & tdef.Name
    End If
Next
outFile.WriteLine ""

' リレーションシップ
outFile.WriteLine "■■■ リレーションシップ ■■■"
outFile.WriteLine ""
For Each rel In db.Relations
    outFile.WriteLine "  " & rel.Table & " → " & rel.ForeignTable
Next
outFile.WriteLine ""

' クエリ一覧とSQL
outFile.WriteLine "■■■ クエリ一覧とSQL定義 ■■■"
outFile.WriteLine ""

queryCount = 0
For Each qdef In db.QueryDefs
    If Left(qdef.Name, 1) <> "~" Then
        queryCount = queryCount + 1
        outFile.WriteLine "【" & queryCount & "】 " & qdef.Name
        outFile.WriteLine String(50, "-")
        outFile.WriteLine qdef.SQL
        outFile.WriteLine ""
        outFile.WriteLine ""
    End If
Next

outFile.WriteLine "=============================================="
outFile.WriteLine "総クエリ数: " & queryCount
outFile.WriteLine "=============================================="

outFile.Close
db.Close

WScript.Echo "出力完了！"
WScript.Echo "クエリ数: " & queryCount
WScript.Echo "出力先: " & outputPath
