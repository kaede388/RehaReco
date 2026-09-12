VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPatientEpisodes"
' clsPatientEpisodes.cls
Option Explicit
Implements IBase

' --- Delegation ---
Private mBase As clsBase

' --- Private Fields ---
Private pEpisodeUUID As String
Private pPatientID As String
Private pStartDate As Date
Private pEndDate As Date
Private pVisitTypeFlag As Long

Private Sub Class_Initialize()
    Set mBase = New clsBase
    Set mBase.mAudit = New clsAuditInfo
    mBase.DeleteFlag = 0
End Sub

' --- Properties ---
Public Property Get RecordID() As Long: RecordID = mBase.RecordID: End Property
Public Property Let RecordID(ByVal Value As Long): mBase.RecordID = Value: End Property

Public Property Get EpisodeUUID() As String: EpisodeUUID = pEpisodeUUID: End Property
Public Property Let EpisodeUUID(ByVal Value As String): pEpisodeUUID = Value: End Property

Public Property Get PatientID() As String: PatientID = pPatientID: End Property
Public Property Let PatientID(ByVal Value As String): pPatientID = Value: End Property

Public Property Get StartDate() As Date: StartDate = pStartDate: End Property
Public Property Let StartDate(ByVal Value As Date): pStartDate = Value: End Property

Public Property Get EndDate() As Date: EndDate = pEndDate: End Property
Public Property Let EndDate(ByVal Value As Date): pEndDate = Value: End Property

Public Property Get VisitTypeFlag() As Long: VisitTypeFlag = pVisitTypeFlag: End Property
Public Property Let VisitTypeFlag(ByVal Value As Long): pVisitTypeFlag = Value: End Property

Public Property Get DeleteFlag() As Long: DeleteFlag = mBase.DeleteFlag: End Property
Public Property Let DeleteFlag(ByVal Value As Long): mBase.DeleteFlag = Value: End Property

Public Property Get mAudit() As clsAuditInfo: Set mAudit = mBase.mAudit: End Property
Public Property Set mAudit(ByVal Value As clsAuditInfo): Set mBase.mAudit = Value: End Property

' --- IBase Implementation ---
Private Property Let IBase_RecordID(ByVal Value As Long): RecordID = Value: End Property
Private Property Get IBase_RecordID() As Long: IBase_RecordID = RecordID: End Property

Private Property Let IBase_DeleteFlag(ByVal Value As Long): DeleteFlag = Value: End Property
Private Property Get IBase_DeleteFlag() As Long: IBase_DeleteFlag = DeleteFlag: End Property

Private Property Set IBase_mAudit(ByVal Value As clsAuditInfo): Set mAudit = Value: End Property
Private Property Get IBase_mAudit() As clsAuditInfo: Set IBase_mAudit = mAudit: End Property

Private Property Get IBase_IsValid() As Boolean: IBase_IsValid = Validate And mBase.IsValid: End Property

' --- Validate ---
Public Function Validate() As Boolean
    ' バリデーションチェック
    If Trim(pEpisodeUUID) = "" Or Len(pEpisodeUUID) <> 36 Then
        Validate = False
        Exit Function
    End If
    
    If Not IsNumeric(pPatientID) Or Len(pPatientID) > 8 Then
        Validate = False
        Exit Function
    End If
    
    If pStartDate > pEndDate Then
        Validate = False
        Exit Function
    End If
    
    Validate = True
End Function