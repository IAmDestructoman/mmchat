VERSION 5.00
Begin VB.UserControl Socket 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   480
   Windowless      =   -1  'True
End
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event OnClose()
Public Event OnConnect()
Public Event OnConnectionRequest(ByVal requestId As Long)
Public Event OnDataArrival(ByVal bytesTotal As Long)
Public Event OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event OnSendComplete()
Public Event OSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)

Private WithEvents CSck As CSocket
Attribute CSck.VB_VarHelpID = -1


'Public Properties
Public Property Get LocalPort() As Long
    LocalPort = CSck.LocalPort
End Property

Public Property Let LocalPort(NewValue As Long)
    CSck.LocalPort = NewValue
    
End Property

Public Property Get LocalIP() As String
    LocalIP = CSck.LocalIP
End Property

Public Property Get LocalHostName() As String
    LocalHostName = CSck.LocalHostName
End Property

Public Property Get BytesReceived() As Long
    BytesReceived = CSck.BytesReceived
End Property

Public Property Get Protocol() As ProtocolConstants
    Protocol = CSck.Protocol
End Property

Public Property Let Protocol(NewValue As ProtocolConstants)
    CSck.Protocol = NewValue
End Property

Public Property Get RemoteHostIP() As String
    RemoteHostIP = CSck.RemoteHostIP
End Property

Public Property Get RemoteHost() As String
    RemoteHost = CSck.RemoteHost
End Property

Public Property Let RemoteHost(NewValue As String)
    CSck.RemoteHost = newval
End Property

Public Property Get State() As StateConstants
    State = CSck.State
End Property

Public Property Get SocketHandle() As Long
    SocketHandle = CSck.SocketHandle
End Property

Public Property Get RemotePort() As Long
    RemotePort = CSck.RemotePort
End Property

Public Property Let RemotePort(NewValue As Long)
    CSck.RemotePort = NewValue
End Property

'Event Handlers
Private Sub CSck_OnClose()
    RaiseEvent OnClose
End Sub

Private Sub CSck_OnConnect()
    RaiseEvent OnConnect
End Sub

Private Sub CSck_OnConnectionRequest(ByVal requestId As Long)
    RaiseEvent OnConnectionRequest(requestId)
End Sub

Private Sub CSck_OnDataArrival(ByVal bytesTotal As Long)
    RaiseEvent OnDataArrival(bytesTotal)
End Sub

Private Sub CSck_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent OnError(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub CSck_OnSendComplete()
    RaiseEvent OnSendComplete
End Sub

Private Sub CSck_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    RaiseEvent OSendProgress(bytesSent, bytesRemaining)
End Sub

Private Sub UserControl_Initialize()
    Set CSck = New CSocket
End Sub

'Public Subs & Functions
Public Sub Accept(requestId As Long)
    CSck.Accept requestId
End Sub

Public Sub Bind(Optional lngLocalPort As Long, Optional strLocalIP As String)
    CSck.Bind lngLocalPort, strLocalIP
End Sub

Public Sub CloseSocket()
    CSck.CloseSocket
End Sub
    
Public Sub Connect(Optional strRemoteHost As Variant, Optional lngRemotePort As Variant)
    CSck.Connect strRemoteHost, lngRemotePort
End Sub

Public Sub GetData(varData As Variant, Optional varType As Variant, Optional maxLen As Variant)
    CSck.GetData varData, varType, maxLen
End Sub

Public Sub Listen()
    CSck.Listen
End Sub

Public Sub PeekData(varData As Variant, Optional varType As Variant, Optional maxLen As Variant)
    CSck.PeekData varData, varType, maxLen
End Sub

Public Sub SendData(varData As Variant)
    CSck.SendData varData
End Sub
Public Function vbSocket() As Long
    vbSocket = CSck.vbSocket
End Function
    
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    CSck.LocalPort = PropBag.ReadProperty("LocalPort", CSck.LocalPort)
    CSck.RemoteHost = PropBag.ReadProperty("RemoteHost", CSck.RemoteHost)
    CSck.RemotePort = PropBag.ReadProperty("RemotePort", CSck.RemotePort)
    CSck.Protocol = PropBag.ReadProperty("Protocol", CSck.Protocol)
End Sub

Private Sub UserControl_Resize()
    Width = 32 * Screen.TwipsPerPixelX
    Height = 32 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "LocalPort", CSck.LocalPort
    PropBag.WriteProperty "RemoteHost", CSck.RemoteHost
    PropBag.WriteProperty "RemotePort", CSck.RemotePort
    PropBag.WriteProperty "Protocol", CSck.Protocol
End Sub
