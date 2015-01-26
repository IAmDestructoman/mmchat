Attribute VB_Name = "Module1"

Public abold As String
Public areset As String
Public ablack As String
Public ablue As String
Public acyan As String
Public agreen As String
Public amagenta As String
Public ared As String
Public ayellow As String
Public awhite As String

Public sysuptime As String
Public init As Boolean
Public sysversion As String
Public welcome As String
Public datime As Variant
Public authdis As String
Public authserv As String
Public authnum As String
Public authhost As String
Public gaglist(1 To 50, 1 To 2) As Variant
Public Type userinfo
    name As String
    cip As String
    rip As String
    pass As String
    perm As Integer
    sperm As Integer
    ver As String
    ocn As String
    contime As String
    lastchattime As String
    idletime As Variant
    pingtime As String
    gtimer As String
    gaggedtime As Variant
    log() As String
    channel As Integer
    idle As Boolean
End Type
Public Type channelinfo
    name As String
    password As String
    permlvl As String
    authmode As String
    log() As String
    sil As Boolean
End Type
Public Type commandinfo
    name As String
    permlvl As String
    sparams As Integer
    nparams As Integer
    help As String
End Type
Public procplugin As Boolean
Public channels() As channelinfo
Global users() As userinfo
Public plugins() As Object
Public commands(1 To 27) As commandinfo
Public MOTD As String
Public isservice As Boolean


Sub timeout(X)
start = Timer
Do: DoEvents
Loop While start + X > Timer
End Sub

