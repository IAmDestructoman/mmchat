VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox List3 
      Height          =   3180
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3120
      Top             =   3840
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3480
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dafile
Private pa(1 To 2000, 1 To 2000) As Byte
Private timer1count As Byte

Private Sub Form_Load()
On Error Resume Next
abold = Chr(27) & "[1m"
areset = Chr(27) & "[0m"
ablack = Chr(27) & "[30m"
ablue = Chr(27) & "[34m"
acyan = Chr(27) & "[36m"
agreen = Chr(27) & "[32m"
amagenta = Chr(27) & "[35m"
ared = Chr(27) & "[31m"
ayellow = Chr(27) & "[33m"
awhite = Chr(27) & "[37m"

timer1count = 1
picx = 1
picy = 1
If Dir(App.Path & "\hololist.txt", vbNormal) <> "" Then
    Open App.Path & "\hololist.txt" For Input As 1
    dafile = Input(LOF(1), 1)
    Close 1
    Dim x() As String
    x() = Split(dafile, vbCrLf)
    For a = 0 To UBound(x)
    If x(a) <> "" Then List3.AddItem x(a)
    Next a

End If

Winsock1.Connect "www.medievia.com", 80

'doit
End Sub

Private Sub doit()
On Error Resume Next
daloc = 0

ReDim dadata(Len(dafile))
For x = 0 To Len(dafile) - 1
dadata(x) = Asc(Mid(dafile, x + 1, 1))
Next x

'Open "h:\redirect\desktop\livemapinfo2.gif" For Binary Access Read As 1
'ReDim dadata(LOF(1))
'Get 1, , dadata

For a = 0 To 5
x = x & Chr(GN)
Next a
'screen desc
'screen width
x = Val(GN + 256 * GN)
'pic.Width = x + 1

'screen height
x = Val(GN + 256 * GN)

'pic.Height = x + 1

'byte 5 of S. D.
x = Val(GN)
gctsize = 2 ^ (bin2dec(getbit(x, 2) & getbit(x, 1) & getbit(x, 0)) + 1)

'byte 6 background
x = Val(GN)

'byte 7 all 0
x = Val(GN)


'build Global color table.
ReDim gct(0 To gctsize - 1)
For a = 0 To gctsize - 1
gct(a) = RGB(GN, GN, GN)
Next a
'end global color table
'skip to image descriptor
Do: DoEvents
Loop While GN <> 44
imgleft = GN
imgleft = imgleft + 256 * Val(GN)

imgtop = GN
imgtop = imgtop + 256 * Val(GN)

imgwidth = GN
imgwidth = imgwidth + 256 * Val(GN)
imgheight = GN
imgheight = imgheight + 256 * Val(GN)
x = GN 'misc data. not going to use for now.
mcs = GN ' get min code size = 5 clear code = 32 eoi = 33
ccs = mcs + 1
x = GN ' how many bytes in this sub block
ReDim codetab(0 To (2 ^ mcs) - 1) As String
Dim block As String
Dim kk() As String

Do

For temp = 1 To x
block = bytetobin(GN) & block
'If Len(block) < ccs Then block = bytetobin(GN) & block
Next temp
xxx = Len(block)
'For temp = 1 To x
    
   
    
    Do: DoEvents
        
        If Len(block) >= ccs Then
            dacode = bin2dec(Right$(block, ccs))
            
            block = Mid(block, 1, Len(block) - ccs)
        End If
        'Form1.Caption = xxx - Len(block) & " of " & xxx
        Select Case dacode
        Case (2 ^ mcs)
            ReDim Preserve codetab(0 To (2 ^ mcs) + 1)
            For temp2 = 0 To (2 ^ mcs) - 1
            codetab(temp2) = Str(temp2)
            Next temp2
            ccs = mcs + 1
            dacodem = vbNullString
        Case (2 ^ mcs) + 1
            
            Close 1
            
            Exit Sub
        Case Else
            If dacode > UBound(codetab) Then
                kk() = Split(codetab(dacodem), ",")
                k = kk(0)
                drawit codetab(dacodem) & "," & k
                
                ReDim Preserve codetab(0 To UBound(codetab) + 1)
                codetab(dacode) = codetab(dacodem) & "," & k
                'List1.AddItem dacode & "-" & codetab(dacode)
               
                'List1.ListIndex = List1.ListCount - 1
                
                dacodem = dacode
                If (2 ^ ccs) - 1 <= UBound(codetab) And ccs < 12 Then
                ccs = ccs + 1
                End If
            Else
                drawit codetab(dacode)
                
                kk() = Split(codetab(dacode), ",")
                k = kk(0)
                If dacodem <> vbNullString Then
                    ReDim Preserve codetab(0 To UBound(codetab) + 1)
                    codetab(UBound(codetab)) = codetab(dacodem) & "," & k
                    'List1.AddItem UBound(codetab) & "-" & codetab(dacodem) & "," & k
                    'List1.ListIndex = List1.ListCount - 1
                    If (2 ^ ccs) - 1 <= UBound(codetab) And ccs < 12 Then
                    ccs = ccs + 1
                    End If
                End If
                dacodem = dacode
            End If
                        
        End Select
    Loop While Len(block) > ccs - 1
'Next temp
x = GN ' get next block count
Loop While x > 0


End Sub
Sub findthem()
On Error Resume Next
List1.Clear
For a = 4 To 1996
For b = 4 To 1996
    DoEvents
    
    If a = 1001 And b = 1063 Then
        a = a
    End If
    If pa(a, b) = 4 Then
        
        isholo = True
        If pa(a - 1, b - 1) = 4 Or pa(a - 1, b + 1) = 4 Or pa(a + 1, b - 1) = 4 Or pa(a + 1, b + 1) = 4 Then isholo = False
        c = 0
        If pa(a - 1, b) = 4 Then c = c + 1
        If pa(a + 1, b) = 4 Then c = c + 1
        If pa(a, b + 1) = 4 Then c = c + 1
        If pa(a, b - 1) = 4 Then c = c + 1
        If c <> 3 Then
            isholo = False
        End If
        c = 0
        If pa(a - 2, b - 2) = 4 Then c = c + 1
        If pa(a + 2, b - 2) = 4 Then c = c + 1
        If pa(a - 2, b + 2) = 4 Then c = c + 1
        If pa(a + 2, b + 2) = 4 Then c = c + 1
        If c = 4 Then
            isholo = False
        End If
        If isholo Then
            road = 99
            For aa = a - 5 To a + 5
            For bb = b - 5 To b + 5
                If gct(pa(aa, bb)) = RGB(255, 0, 255) Then
                    c = Abs(aa - a) + Abs(bb - b)
                    If c < road Then road = c
                End If
            Next bb
            Next aa

            List1.AddItem a & "," & b & "," & road
        End If
    End If
Next b
Next a
End Sub
Sub comparethem()
On Error Resume Next
'check for 1st run
If List2.ListCount = 0 Then
    For a = 0 To List1.ListCount - 1
    List2.AddItem List1.List(a)
    Next a
End If
'remove
For a = List3.ListCount - 1 To 0 Step -1
    inlist = False
    c = InStr(List3.List(a), ",")
    c = InStr(c + 1, List3.List(a), ",")
    c = Mid(List3.List(a), 1, InStr(c + 1, List3.List(a), ",") - 1)
    
    For b = 0 To List1.ListCount - 1

        
        If c = List1.List(b) Then inlist = True
    Next b
    If inlist = False Then
        List3.RemoveItem a
    End If
    
Next a
'add
didupdate = False
c = List3.ListCount
For a = 0 To List1.ListCount - 1
    inlist = False
    For b = 0 To List2.ListCount - 1
        If List1.List(a) = List2.List(b) Then inlist = True
    Next b
    If inlist = False Then
        List3.AddItem List1.List(a) & "," & Now
        List3.ListIndex = List3.ListCount - 1
        'format$("asdf","!@@@@@@@@@")
        didupdate = True
    End If
Next a
'If List3.ListCount > c + 10 Then
'    For a = c To List3.ListCount - 1
'    List3.RemoveItem (c)
'    Next a
'    didupdate = False
'End If
If didupdate = True Then sendupdate

'copy list1 to list2
List2.Clear
For a = 0 To List1.ListCount - 1
    List2.AddItem List1.List(a)
Next a
End Sub
Sub sendupdate()
On Error Resume Next
'daage = "Age:       NEW |       NEW | "
'daloc = "Loc: 1539  549 | 1539  549 | "
'dafly = "Fly: E 54 N 45 | E 54 N 45 | "
Dim k() As String

datext = Chr(10) & abold
For a = List3.ListCount - 1 To c Step -1
    k() = Split(List3.List(a), ",")
    If DateDiff("n", k(3), Now) > 240 Then
        List3.RemoveItem a
    End If
    If DateDiff("n", k(3), Now) = 0 Then
        k(3) = "NEW("
        datext = datext & acyan & "NEW" & ablue & "("
    Else
        datext = datext & acyan & DateDiff("n", k(3), Now) & awhite & "Mins" & ablue & "("
    End If
        If k(0) > 1000 Then
        ew = "E"
        k(0) = Round((k(0) - 1000) / 10, 0)
    Else
        ew = "W"
        k(0) = Round((1000 - k(0)) / 10, 0)
    End If
    If k(1) > 1000 Then
        ns = " S"
        k(1) = Round((k(1) - 1000) / 10, 0)
    Else
        ns = " N"
        k(1) = Round((1000 - k(1)) / 10, 0)
    End If
    datext = datext & ayellow & ew & awhite & k(0) & ayellow & ns & awhite & k(1) & ablue & ")"
    If k(2) <> 99 Then
        datext = datext & ared & "Road:" & k(2)
    End If
    datext = datext & "  "
Next a

'datext = Chr(10) & daage & Chr(10) & daloc2 & Chr(10) & dafly & Chr(10)

If Dir(App.Path & "\hololist.txt", vbNormal) <> "" Then Kill App.Path & "\hololist.txt"
If List3.ListCount = 0 Then Exit Sub
Open App.Path & "\hololist.txt" For Output As 1
For a = 0 To List3.ListCount - 1
Print #1, List3.List(a)
Next a
Close 1

daupdate = "HoloFind: (Age - Fly Dirs)" & datext


End Sub


Sub drawit(ByVal asdf As String)
On Error Resume Next
asdf = Replace(asdf, " ", vbNullString)
If InStr(asdf, ",") = 0 Then
    'pic.PSet (picx, picy), gct(Val(asdf))
    pa(picx, picy) = Val(asdf)

    If picx = imgwidth Then
        picx = 1
        picy = picy + 1
    Else
        picx = picx + 1
    End If
Else
    Dim ww() As String
    ww() = Split(asdf, ",")
    For a = 0 To UBound(ww)
    w = ww(a)
        'pic.PSet (picx, picy), gct(Val(w))
        pa(picx, picy) = Val(w)

        If picx = imgwidth Then
            picx = 1
            picy = picy + 1
        Else
            picx = picx + 1
        End If
    Next a
    
End If

End Sub
Public Function bin2dec(Binary As String) As Long
On Error Resume Next
Dim n As Long
Dim s As Integer
Binary = Replace(Binary, " ", vbNullString)
    For s = 1 To Len(Binary)
        n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ _
            (s - 1)))
    Next s

    bin2dec = n
End Function


Function GN() As Byte
On Error Resume Next
GN = dadata(daloc)
daloc = daloc + 1
End Function
Function getbit(ByVal dabyte As Byte, whatbit As Integer) As Byte
On Error Resume Next
Dim bytetobit(7) As String
bytetobit(7) = "0"
bytetobit(6) = "0"
bytetobit(5) = "0"
bytetobit(4) = "0"
bytetobit(3) = "0"
bytetobit(2) = "0"
bytetobit(1) = "0"
bytetobit(0) = "0"
If dabyte - 128 >= 0 Then
    bytetobit(7) = "1"
    dabyte = dabyte - 128
End If
If dabyte - 64 >= 0 Then
    bytetobit(6) = "1"
    dabyte = dabyte - 64
End If
If dabyte - 32 >= 0 Then
    bytetobit(5) = "1"
    dabyte = dabyte - 32
End If
If dabyte - 16 >= 0 Then
    bytetobit(4) = "1"
    dabyte = dabyte - 16
End If
If dabyte - 8 >= 0 Then
    bytetobit(3) = "1"
    dabyte = dabyte - 8
End If
If dabyte - 4 >= 0 Then
    bytetobit(2) = "1"
    dabyte = dabyte - 4
End If
If dabyte - 2 >= 0 Then
    bytetobit(1) = "1"
    dabyte = dabyte - 2
End If
If dabyte - 1 >= 0 Then
    bytetobit(0) = "1"
    dabyte = dabyte - 1
End If
Select Case whatbit
Case 0
    getbit = bytetobit(0)
Case 1
    getbit = bytetobit(1)
Case 2
    getbit = bytetobit(2)
Case 3
    getbit = bytetobit(3)
Case 4
    getbit = bytetobit(4)
Case 5
    getbit = bytetobit(5)
Case 6
    getbit = bytetobit(6)
Case 7
    getbit = bytetobit(7)
End Select

End Function


Public Function bytetobin(DecimalNum As Byte) As String
On Error Resume Next
Dim tmp As String
Dim n As Long

    n = DecimalNum

    tmp = Trim(Str(n Mod 2))
    n = n \ 2

    Do While n <> 0
        tmp = Trim(Str(n Mod 2)) & tmp
        n = n \ 2
    Loop
For x = 1 To 8 - Len(tmp)
tmp = "0" & tmp
Next x
    bytetobin = tmp
    
End Function


Private Sub Timer1_Timer()
On Error Resume Next
Form1.Caption = Now
timer1count = timer1count + 1
If timer1count < 10 Then Exit Sub
timer1count = 0
picx = 1
picy = 1
Winsock1.Connect "www.medievia.com", 80
Form1.Caption = "connecting... " & Now
End Sub

Private Sub winsock1_Close()
On Error Resume Next
crlfpos = InStr(dafile, vbCrLf & vbCrLf) + 4
dafile = Mid(dafile, crlfpos, Len(dafile))
Winsock1.Close
doit
findthem
comparethem

End Sub

Private Sub winsock1_Connect()
On Error Resume Next
Winsock1.SendData "GET /images/livemaps/livemapinfo.gif HTTP/1.1" & vbCrLf & "Host: www.medievia.com" & vbCrLf & "Connection: close" & vbCrLf & vbCrLf
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData dadata2, vbString
dafile = dafile & dadata2
End Sub
Private Sub winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close

End Sub

Sub timeout(x)
xx = timer
Do: DoEvents
Loop While timer < xx + x
End Sub






