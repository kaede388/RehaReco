VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBase"

' 全てのエンティティクラスが実装すべき共通プロパティの定義

Option Explicit

' レコードID
Public Property Get RecordID() As Long: End Property
Public Property Let RecordID(ByVal Value As Long): End Property

' 削除フラグ
Public Property Get DeleteFlag() As Long: End Property
Public Property Let DeleteFlag(ByVal Value As Long): End Property

' 監査情報
Public Property Get mAudit() As clsAuditInfo: End Property
Public Property Set mAudit(ByVal Value As clsAuditInfo): End Property

' 有効性チェック
Public Property Get IsValid() As Boolean: End Property