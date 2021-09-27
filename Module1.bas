Attribute VB_Name = "Module1"
Option Explicit

'/**
' * UiPathからコールされることを前提としたExcel VBAフレームワーク。
' * 本処理は必ず標準モジュールに対してインポートして使って下さい。
' *
' * 必要に応じて引数を追加する。
' * 実処理自体は本関数ではなく、Main() に記載すること。
' *
' * ※参考：UiPathナレッジベース：VBA Macroの自動化対応時のエラーハンドリング（OnErrorGoto）
' * https://www.uipath.com/ja/resources/knowledge-base/vba-macro-onerrorgoto
' *
' * ※参考：Qiita：タイトル
' * URL
' *
' * @param {型} 変数名 - 1つめの引数
' * @param {型} 変数名 - 2つめの引数
' * @return {String} - 正常時は空白, 異常時はエラー内容を返す
' */
Public Function UiPathVBAFramework() As String
    ' エラートラップ
    On Error GoTo ERR_CATCH_

    Dim Ret As String
    Ret = ""    ' 返却値を初期化

    Application.ScreenUpdating = False  ' 画面更新OFF
    Application.DisplayAlerts = False   ' 警告表示OFF

    ' メイン処理は Main() に記載
    Call Main

' 終了処理
EXIT_:
    Application.ScreenUpdating = False  ' 画面更新ON
    Application.DisplayAlerts = False   ' 警告表示ON

    UiPathVBAFramework = Ret

    On Error GoTo 0
    Exit Function

' --------------------------
' ----- エラーキャッチ -----
' --------------------------
ERR_CATCH_:
    ' エラー内容を返却値に保持
    Ret = "Excelマクロ内でエラーが発生しました。" & vbCrLf & _
          "エラー番号：" & Err.Number & vbCrLf & _
          "エラー内容：" & Err.Description
    Err.Clear

    ' 終了処理へ
    Resume EXIT_
End Function

'/**
' * 実処理を記載する関数。
' * 必要に応じて引数を追加する。
' *
' * @param {型} 変数名 - 1つめの引数
' * @param {型} 変数名 - 2つめの引数
' */
Private Sub Main()
    ' 実処理を記載
End Sub
