VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPatientInfoRepo"
Option Explicit

Public Sub Save(patientInfo As clsPatientInfo)
    On Error GoTo ErrorHandler
    Dim CN As ADODB.Connection
    Dim SQL As String
    Set CN = New ADODB.Connection
    CN.Open GetConnectionString()
    
    ' INSERT ... ON DUPLICATE KEY UPDATE による Upsert
    SQL = "INSERT INTO Patient_Information " & _
          "(PatientID, Name1, Name2, Name3, Name4, Sex, DateOfBirth, DeleteFlag, " & _
          "InputDateTime, InputStaffID, InputPC, UpdateDateTime, UpdateStaffID, UpdatePC) " & _
          "VALUES ('" & patientInfo.PatientID & "', '" & patientInfo.Name1 & "', '" & patientInfo.Name2 & "', '" & patientInfo.Name3 & "', '" & patientInfo.Name4 & "', " & _
          patientInfo.Sex & ", '" & Format(patientInfo.DateOfBirth, "yyyy-mm-dd") & "', " & patientInfo.DeleteFlag & ", " & _
          "Now(), '" & patientInfo.Audit.InputStaffID & "', '" & patientInfo.Audit.InputPC & "', " & _
          "Now(), '" & patientInfo.Audit.InputStaffID & "', '" & patientInfo.Audit.InputPC & "') " & _
          "ON DUPLICATE KEY UPDATE " & _
          "Name1='" & patientInfo.Name1 & "', Name2='" & patientInfo.Name2 & "', Name3='" & patientInfo.Name3 & "', Name4='" & patientInfo.Name4 & "', " & _
          "Sex=" & patientInfo.Sex & ", DateOfBirth='" & Format(patientInfo.DateOfBirth, "yyyy-mm-dd") & "', " & _
          "DeleteFlag=" & patientInfo.DeleteFlag & ", UpdateDateTime=Now(), " & _
          "UpdateStaffID='" & patientInfo.Audit.UpdateStaffID & "', UpdatePC='" & patientInfo.Audit.UpdatePC & "'"
    
    CN.Execute SQL
    CN.Close
    Set CN = Nothing
    Exit Sub
ErrorHandler:
    Call modErrorManager.HandleError("clsPatientInfoRepo", "Save", Err.Number, Err.Description)
    If Not CN Is Nothing Then If CN.State <> 0 Then CN.Close: Set CN = Nothing
End Sub

Private Function CalculateAge(ByVal DateOfBirth As Date, ByVal BaseDate As Date) As Long
    On Error GoTo ErrorHandler
    If DatePart("m", DateOfBirth) > DatePart("m", BaseDate) Or (DatePart("m", DateOfBirth) = DatePart("m", BaseDate) And DatePart("d", DateOfBirth) > DatePart("d", BaseDate)) Then
        CalculateAge = DateDiff("yyyy", DateOfBirth, BaseDate) - 1
    Else
        CalculateAge = DateDiff("yyyy", DateOfBirth, BaseDate)
    End If
    Exit Function

ErrorHandler:
    CalculateAge = 0
    Call modErrorManager.HandleError("clsPatientInfoRepo", "CalculateAge", Err.Number, Err.Description)
End Function

' 定IDの患者報を取
Public Function GetLoadByPatientID(ByVal patientID As String) As clsPatientInfo
    On Error GoTo ErrorHandler
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim SQL As String
    Dim Patient As clsPatientInfo

    Set CN = New ADODB.Connection
    CN.ConnectionString = GetConnectionString()
    CN.Open

    SQL = "SELECT * FROM Patient_Information WHERE PatientID = '" & patientID & "' AND DeleteFlag = 0 LIMIT 1"
    Set RS = CN.Execute(SQL)

    If Not RS.EOF Then
        Set Patient = New clsPatientInfo
        Patient.PatientID = RS("PatientID").Value
        Patient.Name1 = RS("Name1").Value
        Patient.Name2 = RS("Name2").Value
        Patient.Name3 = RS("Name3").Value
        Patient.Name4 = RS("Name4").Value
        Patient.Sex = RS("Sex").Value
        If Not IsNull(RS("DateOfBirth").Value) Then Patient.DateOfBirth = RS("DateOfBirth").Value
        Patient.Age = CalculateAge(Patient.DateOfBirth, Date)
    Else
        Set Patient = Nothing
    End If

    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing

    Set GetLoadByPatientID = Patient
    Exit Function
ErrorHandler:
    Call modErrorManager.HandleError("clsPatientInfoRepo", "GetLoadByPatientID", Err.Number, Err.Description)
    If Not RS Is Nothing Then If RS.State <> 0 Then RS.Close: Set RS = Nothing
    If Not CN Is Nothing Then If CN.State <> 0 Then CN.Close: Set CN = Nothing
    Set GetLoadByPatientID = Nothing
End Function

' 患者を論理削除
Public Sub Delete(ByVal patientID As String, ByVal staffID As String, ByVal pcName As String)
    On Error GoTo ErrorHandler
    Dim CN As ADODB.Connection
    Dim SQL As String
    Set CN = New ADODB.Connection
    CN.ConnectionString = GetConnectionString()
    CN.Open
    
    SQL = "UPDATE Patient_Information SET " & _
          "DeleteFlag = 1, " & _
          "UpdateDateTime = Now(), " & _
          "UpdateStaffID = '" & staffID & "', " & _
          "UpdatePC = '" & pcName & "' " & _
          "WHERE PatientID = '" & patientID & "'"
    CN.Execute SQL
    
    CN.Close
    Set CN = Nothing
    Exit Sub
ErrorHandler:
    Call modErrorManager.HandleError("clsPatientInfoRepo", "Delete", Err.Number, Err.Description)
    If Not CN Is Nothing Then If CN.State <> 0 Then CN.Close: Set CN = Nothing
End Sub

' 全患者を取得
Public Function GetAllPatients() As Collection
    On Error GoTo ErrorHandler
    Dim CN As ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim SQL As String
    Dim p As clsPatientInfo
    Dim patients As Collection
    
    Set patients = New Collection
    Set CN = New ADODB.Connection
    CN.ConnectionString = GetConnectionString()
    CN.Open
    
    SQL = "SELECT * FROM Patient_Information WHERE DeleteFlag = 0 ORDER BY PatientID"
    Set RS = CN.Execute(SQL)
    
    Do While Not RS.EOF
        Set p = New clsPatientInfo
        p.PatientID = RS("PatientID").Value
        p.Name1 = RS("Name1").Value
        p.Name2 = RS("Name2").Value
        p.Name3 = RS("Name3").Value
        p.Name4 = RS("Name4").Value
        p.Sex = RS("Sex").Value
        If Not IsNull(RS("DateOfBirth").Value) Then p.DateOfBirth = RS("DateOfBirth").Value
        
        patients.Add p
        RS.MoveNext
    Loop
    
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    Set GetAllPatients = patients
    Exit Function
ErrorHandler:
    Call modErrorManager.HandleError("clsPatientInfoRepo", "GetAllPatients", Err.Number, Err.Description)
    If Not RS Is Nothing Then If RS.State <> 0 Then RS.Close: Set RS = Nothing
    If Not CN Is Nothing Then If CN.State <> 0 Then CN.Close: Set CN = Nothing
    Set GetAllPatients = New Collection
End Function