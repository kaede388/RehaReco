VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEpisodeRepo"
' clsEpisodeRepo.bas
' 患者エピソード（入院・外来単位）を管理するリポジトリ Class Module版
' 依存テーブル: patient_episodes, patient_information
Option Explicit

Private Const SourceModule As String = "clsEpisodeRepo"

' 指定日付に入院中の患者リストを取得するメソッド
Public Function GetActivePatients(targetDate As Date) As Collection
    On Error GoTo ErrHandler
    Dim Coll As New Collection
    Dim CN As ADODB.Connection: Set CN = New ADODB.Connection
    Dim RS As ADODB.Recordset
    Dim SQL As String
    
    CN.Open GetConnectionString()
    
    ' Concatをやめ、個別に取得する
    SQL = "SELECT e.EpisodeUUID, p.PatientID, p.Name1, p.Name2 " & _
          "FROM patient_information p " & _
          "INNER JOIN patient_episodes e ON p.PatientID = e.PatientID " & _
          "WHERE e.StartDate <= '" & Format(targetDate, "yyyy-mm-dd") & "' " & _
          "AND (e.EndDate IS NULL OR e.EndDate >= '" & Format(targetDate, "yyyy-mm-dd") & "')"
    
    Set RS = CN.Execute(SQL)
    
    Dim Item As clsResponsibleStaff
    
    Do While Not RS.EOF
        Set item = New clsResponsibleStaff
        item.EpisodeUUID = RS("EpisodeUUID").Value
        item.PatientID = RS("PatientID").Value
        item.PatientName1 = RS("Name1").Value
        item.PatientName2 = RS("Name2").Value
        Coll.Add item
        RS.MoveNext
    Loop
    
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    Set GetActivePatients = Coll
    
    Exit Function

ErrHandler:
    If Not RS Is Nothing Then If RS.State = adStateOpen Then RS.Close: Set RS = Nothing
    If Not CN Is Nothing Then If CN.State = adStateOpen Then CN.Close: Set CN = Nothing
    Call HandleError(SourceModule, "GetActivePatients", Err.Number, Err.Description)
End Function

' --- 取得または新規作成：EpisodeUUID ---
Public Function GetOrCreateEpisodeUUID(patientID As String, startDate As Date, staffID As String) As String
    Dim RS As ADODB.Recordset, SQL As String, newUUID As String
    Dim CN As ADODB.Connection
    
    On Error GoTo ErrHandler
    
    Set CN = New ADODB.Connection
    CN.ConnectionString = GetConnectionString()
    CN.Open
    
    SQL = "SELECT EpisodeUUID " & _
            "FROM patient_episodes " & _
            "WHERE PatientID='" & patientID & "' " & _
            "AND StartDate='" & Format(startDate, "yyyy-mm-dd") & "' " & _
            "LIMIT 1"
    Set RS = CN.Execute(SQL)
    
    If Not RS.EOF Then
        GetOrCreateEpisodeUUID = RS("EpisodeUUID").Value
    Else
        newUUID = GetUUID() 
        SQL = "INSERT INTO patient_episodes (EpisodeUUID, PatientID, StartDate, VisitTypeFlag, InputDateTime, InputStaffID, InputPC, UpdateDateTime, UpdateStaffID, UpdatePC) " & _
                "VALUES ('" & newUUID & "', '" & patientID & "', '" & Format(startDate, "yyyy-mm-dd") & "', 1, Now(), '" & staffID & "', '" & Environ("COMPUTERNAME") & "', Now(), '" & staffID & "', '" & Environ("COMPUTERNAME") & "')"
        CN.Execute SQL
        GetOrCreateEpisodeUUID = newUUID
    End If
    
    RS.Close: Set RS = Nothing
    CN.Close: Set CN = Nothing
    
    Exit Function

ErrHandler:
    If Not RS Is Nothing Then If RS.State = adStateOpen Then RS.Close: Set RS = Nothing
    If Not CN Is Nothing Then If CN.State = adStateOpen Then CN.Close: Set CN = Nothing
    Call HandleError(SourceModule, "GetOrCreateEpisodeUUID", Err.Number, Err.Description)
End Function