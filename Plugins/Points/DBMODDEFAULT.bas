Attribute VB_Name = "DBMODDEFAULT"
Option Explicit
'declare all tables in the db
Global DBZoneList As dao.Recordset
Global DBPtsList As dao.Recordset

Sub RefreshSQL()
   Dim strSQL As String

   strSQL = "SELECT * FROM [zonelist]"
   Set DBZoneList = db.OpenRecordset(strSQL)

   strSQL = "SELECT * FROM [ptslist]"
   Set DBPtsList = db.OpenRecordset(strSQL)
End Sub
