VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Secured Intranet Messaging System - Osazuwa Dumnodi Henrietta [SCN0308563]"
   ClientHeight    =   5955
   ClientLeft      =   7635
   ClientTop       =   4170
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmServerTimer 
      Interval        =   100
      Left            =   5760
      Top             =   4440
   End
   Begin MSWinsockLib.Winsock sktConnection 
      Index           =   0
      Left            =   5280
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtTotalIncoming 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
      Width           =   6615
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      Caption         =   "&Shut Down Server"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00004080&
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtServerLog 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1320
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   525
      Index           =   1
      Left            =   3000
      Picture         =   "frmServer.frx":0000
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Osazuwa Dumnodi Henrietta [SCN0308563]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Width           =   6615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "This application was developed by:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ON!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   8
      Top             =   -240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IM Firewall Is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   -240
      Width           =   1575
   End
   Begin VB.Label lblActiveConnections 
      BackColor       =   &H80000009&
      Caption         =   "0 Active Connections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1620
      Index           =   0
      Left            =   5040
      Picture         =   "frmServer.frx":0A35
      Top             =   -120
      Width           =   1800
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "IM SERVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Chat Session"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Server Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gHandles() As String
Private gSentYN() As Boolean
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private gMessages() As String

Private Sub cmdClose_Click()
End
End Sub

Private Sub Form_Load()
      ' Only connection(0) listens, all others are connections
    sktConnection(0).Close
    sktConnection(0).LocalPort = gPort
    sktConnection(0).Listen
    
    ReDim gHandles(1) As String
    gHandles(0) = "IM Server"
    AddToServerLog ("Server is Listening on port " & gPort)
    
    End Sub
Private Sub sktConnection_Close(Index As Integer)
    ' Announce the new departure
    AddToTotalIncoming (gHandles(Index) & " just left")
    
    'close the connection
    sktConnection(Index).Close
    
End Sub

Private Sub sktConnection_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    ' Now accept the new connection
    'A connection was requested from the server.

    Dim i As Integer
    Dim iConnection As Integer

    'Make sure this is control 0 in the array.
    'This is the only one that can accept connections.
    If Index = 0 Then
    
        'Search for available Winsock control.
        For i = 1 To gNumConnections
            If sktConnection(i).State = sckClosed Then
                iConnection = i
                Exit For
            End If
        Next i
        
        'If none was found, create a new one.
        If iConnection = 0 Then
        
            ' Tell the world there's a new connection
            gNumConnections = gNumConnections + 1
            
            'Load a new Winsock control for this connection.
            Load sktConnection(gNumConnections)
            
            ' This connection needs a handle
            ReDim Preserve gHandles(gNumConnections) As String
            ReDim Preserve gSentYN(gNumConnections) As Boolean
            ReDim Preserve gMessages(gNumConnections) As String
            
            ' Catch this user up on the previous conversation
            ' This way they don't get resent the chat session to date
            gMessages(gNumConnections) = gTotalincoming
            
            ' set their handle
            gHandles(gNumConnections) = "unknown"
            
            ' Add to the servers connections
            lblActiveConnections.Caption = gNumConnections & " Active Connections"
            
            'Control to be used is this new control.
            iConnection = gNumConnections
        End If
        
        'Set port for this control to 0.  (Randomly assigns an available port.)
        sktConnection(iConnection).LocalPort = 0
        
        'Have this control accept the connection.
        sktConnection(iConnection).Accept requestID
        
        ' Send the welcome message
        sktConnection(iConnection).SendData "Welcome to the IM System, enter your handle"
    End If
    
End Sub

Private Sub sktConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    'Data has arrived at the server from an open connection.
    Dim newdata As String
        
    'Get the data.
    sktConnection(Index).GetData newdata
    
    If gHandles(Index) = "unknown" Then
    
        ' Store it internally
        gHandles(Index) = newdata
        
        ' Announce the new arrival
        AddToTotalIncoming (newdata & " just joined")
    Else
        'Pass the index of the connection from which the data came.
        AddToTotalIncoming (gHandles(Index) & ":: " & newdata)
    End If
    
End Sub

Private Sub sktConnection_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' this should handle all errors in the server world
    AddToServerLog ("Socket error on #" & Index & " " & Description)
End Sub

Private Sub tmServerTimer_Timer()
    txtTotalIncoming.Text = gTotalincoming
    
    'Send the new data to the group
    Call BroadcastMessage
        
End Sub

Private Sub BroadcastMessage()
    Dim i
    Dim myMessage  As String
    Dim conn_index As Integer
    
    ' Set the default connection index
    conn_index = -1
    
    ' Loop through all connections excluding the servers
    ' Hence we start at 1
    For i = 1 To gNumConnections
        ' if this connection has not had an update then
        If gSentYN(i) = False And sktConnection(i).State = sckConnected Then
            ' Get the message we need to deliver
            myMessage = Left(gTotalincoming, Len(gTotalincoming) - Len(gMessages(i)))
            ' update the message store for this user
            gMessages(i) = gTotalincoming
            ' Get the index of this connection
            conn_index = i
            Exit For
        'Else
         '   gSentYN(i) = True
        End If
    Next i
    
    ' This is so we know to
    ' send the data to everyone
    If conn_index = -1 Then
        For i = 1 To gNumConnections
            gSentYN(i) = False
        Next
    End If
    
    If conn_index > -1 Then
    
        ' check the connection's open
        If sktConnection(conn_index).State = sckConnected Then
            
            ' send the data
            sktConnection(conn_index).SendData myMessage
            
            'signal that we've sent to this user
            gSentYN(conn_index) = True
            
        End If
        
    End If

End Sub

