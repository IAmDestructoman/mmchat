Attribute VB_Name = "Plugins"


Public Sub FindPlugIns(ByVal PlugInPath As String)
    Dim strFileName As String

    strFileName = Dir(PlugInPath & "\*.dll", vbNormal)
    Do While strFileName <> ""
        If PlugInExists(strFileName) Then
            MMChat.List4.AddItem strFileName
        Else
            If PlugInExists(strFileName) Then
                MMChat.List4.AddItem strFileName
            End If
        End If
        strFileName = Dir
    Loop
End Sub
Public Function MakeObjectFromPath(strFileName As String) As String
    MakeObjectFromPath = Replace(strFileName, ".dll", ".Plugin")
End Function
Private Function PlugInExists(strFileName As String) As Boolean
    Dim strCreateString As String
    Dim vDummy As Object
    On Error Resume Next
    strCreateString = MakeObjectFromPath(strFileName)
    Set vDummy = Nothing
    Set vDummy = CreateObject(strCreateString)

    If vDummy Is Nothing Then
        PlugInExists = False
        Call Shell("regsvr32 /u /s " & App.Path & "\plugins\" & strFileName)
        Call Shell("regsvr32 /s " & App.Path & "\plugins\" & strFileName)
    Else
        PlugInExists = True
        Set vDummy = Nothing
    End If
End Function

