VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBase"

' clsBase.cls
' IBase インターフェースを実装するクラスのための委譲用共通クラス
Option Explicit

Private pRecordID As Long
Private pDeleteFlag As Long
Private pAudit As clsAuditInfo

' コンストラクタ
Private Sub Class_Initialize()
    ' 必要に応じて初期化処理を記述
End Sub

' プロパティ定義
Public Property Let RecordID(ByVal Value As Long): pRecordID = Value: End Property
Public Property Get RecordID() As Long: RecordID = pRecordID: End Property

Public Property Let DeleteFlag(ByVal Value As Long): pDeleteFlag = Value: End Property
Public Property Get DeleteFlag() As Long: DeleteFlag = pDeleteFlag: End Property

Public Property Set mAudit(ByVal Value As clsAuditInfo): Set pAudit = Value: End Property
Public Property Get mAudit() As clsAuditInfo: Set mAudit = pAudit: End Property

' 基本的なバリデーション
Public Property Get IsValid() As Boolean
    ' RecordIDが設定されており、Auditが存在することを前提とする
    If pRecordID <= 0 Then
        IsValid = False
        Exit Property
    End If
    If pAudit Is Nothing Then
        IsValid = False
        Exit Property
    End If
    IsValid = True
End Property