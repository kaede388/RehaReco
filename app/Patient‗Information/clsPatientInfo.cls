VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPatientInfo"
' apps/PatientInformation/clsPatientInfo.cls
Option Explicit
Implements IBase

' --- Delegation ---
Private mBase As clsBase

' --- Private Fields ---
Private pPatientID As String
Private pName1 As String
Private pName2 As String
Private pName3 As String
Private pName4 As String
Private pSex As Long
Private pDateOfBirth As Date

Private Sub Class_Initialize()
    Set mBase = New clsBase
    Set mBase.mAudit = New clsAuditInfo
    mBase.DeleteFlag = 0
End Sub

' --- Properties ---
Public Property Get RecordID() As Long: RecordID = mBase.RecordID: End Property
Public Property Let RecordID(ByVal Value As Long): mBase.RecordID = Value: End Property

Public Property Get PatientID() As String: PatientID = pPatientID: End Property
Public Property Let PatientID(ByVal Value As String): pPatientID = Value: End Property

Public Property Get Name1() As String: Name1 = pName1: End Property
Public Property Let Name1(ByVal Value As String): pName1 = Value: End Property

Public Property Get Name2() As String: Name2 = pName2: End Property
Public Property Let Name2(ByVal Value As String): pName2 = Value: End Property

Public Property Get Name3() As String: Name3 = pName3: End Property
Public Property Let Name3(ByVal Value As String): pName3 = Value: End Property

Public Property Get Name4() As String: Name4 = pName4: End Property
Public Property Let Name4(ByVal Value As String): pName4 = Value: End Property

Public Property Get Sex() As Long: Sex = pSex: End Property
Public Property Let Sex(ByVal Value As Long): pSex = Value: End Property

Public Property Get DateOfBirth() As Date: DateOfBirth = pDateOfBirth: End Property
Public Property Let DateOfBirth(ByVal Value As Date): pDateOfBirth = Value: End Property

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

Private Property Get IBase_IsValid() As Boolean: IBase_IsValid = mBase.IsValid: End Property