Attribute VB_Name = "DBMOD"
Option Explicit
Global db As dao.Database
Public Sub LoadDatabase(strPath As String)
   Set db = OpenDatabase(strPath)
   RefreshSQL
End Sub
Public Function AddNewRecord(rstTable As Recordset) As Integer
   Dim intCount As Integer
   Dim intNextRecordID As Integer

   With rstTable
      .Requery
      If .BOF = True And .EOF = True Then
         intNextRecordID = 1
      Else
         .MoveLast
         .MoveFirst
         For intCount = 1 To .RecordCount
            If .Fields("ID").Value <> intCount Then
               intNextRecordID = intCount
               Exit For
            End If
            .MoveNext
         Next intCount
         If intNextRecordID = 0 Then intNextRecordID = .RecordCount + 1
      End If
      .AddNew
      .Fields("ID").Value = intNextRecordID
      .Update
   End With

   AddNewRecord = intNextRecordID
End Function
Public Function DeleteRecord(rstTable As Recordset, intID As Integer) As Integer
   With rstTable
      .Requery
      .FindFirst "ID = " & intID
      If .NoMatch = True Then
         DeleteRecord = -1
      Else
         .Delete
         DeleteRecord = intID
      End If
   End With
End Function
Public Sub DeleteAllRecords(rstTable As Recordset)
   With rstTable
      .Requery

      If .BOF = True And .EOF = True Then Exit Sub

      .MoveFirst
      Do While .EOF = False
         .Delete
         .MoveNext
      Loop
   End With
End Sub
Public Function ChangeField(rstTable As Recordset, intID As Integer, strFieldName As String, varValue As Variant) As Integer
   With rstTable
      .Requery
      .FindFirst "ID = " & intID
      If .NoMatch = True Then
         ChangeField = -1
      Else
         .Edit
         .Fields(strFieldName).Value = varValue
         .Update
         ChangeField = intID
      End If
   End With
End Function

