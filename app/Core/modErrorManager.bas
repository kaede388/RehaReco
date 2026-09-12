Attribute VB_Name = "modErrorManager"
' modErrorManager - エラー管理共通モジュール
Option Explicit

' 独自エラーの列挙体
Public Enum CustomErrors
    Err_PatientID_Invalid = vbObjectError + 1001
    Err_Episode_DateMismatch = vbObjectError + 1101
    Err_Unexpected_Error = vbObjectError + 9999
End Enum

' エラーを発生させるための共通メソッド
Public Sub RaiseError(errCode As CustomErrors, Optional source As String = "AppSystem")
    Dim msg As String
    
    ' 番号に応じてメッセージを返す
    Select Case errCode
        Case Err_PatientID_Invalid
            msg = "患者IDの形式が正しくありません。8桁の数値で指定してください。"
        Case Err_Episode_DateMismatch
            msg = "入院開始日が終了日を過ぎています。"
        Case Else
            msg = "予期せぬエラーが発生しました。"
    End Select
    
    Err.Raise errCode, source, msg
End Sub

' エラーを処理してログを表示・記録する共通メソッド
Public Sub HandleError(ByVal source As String, ByVal procedure As String, ByVal errNumber As Long, ByVal errDescription As String)
    Dim msg As String
    msg = "エラーが発生しました。" & vbCrLf & _
          "場所: " & source & "." & procedure & vbCrLf & _
          "番号: " & errNumber & vbCrLf & _
          "内容: " & errDescription
    
    MsgBox msg, vbCritical, "システムエラー"
    
    ' TODO: 必要に応じてファイルやデータベースへのエラーログ書き込み処理を追加
End Sub