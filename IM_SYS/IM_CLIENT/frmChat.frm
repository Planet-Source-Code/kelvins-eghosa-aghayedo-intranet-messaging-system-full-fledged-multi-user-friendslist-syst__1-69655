VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080FF80&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Window"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLog 
      BackColor       =   &H0080C0FF&
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   5055
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "192.9.200.242"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPort 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4320
      TabIndex        =   3
      Text            =   "1212"
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H80000013&
      Height          =   885
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CommandButton bntExit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton bntSend 
      BackColor       =   &H8000000C&
      Caption         =   "&Send"
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sock1 
      Left            =   4920
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblTo 
      BackColor       =   &H0000FF00&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      Caption         =   "Osazuwa Dumnodi Henrietta"
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
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "This application was developed by:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "[SCN0308563]"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gHandles() As String
Private gSentYN() As Boolean
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private gMessages() As String

Private Sub bntExit_Click()
Unload Me
End Sub
Private Sub useDB(cUser As String)
  Dim str As String
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbUsers.mdb;Persist Security Info=False"
    str = "select * from users where username='" & cUser & "'"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    
    txtIP.Text = rs(2)
    rs.Close
    End Sub
Private Sub bntSend_Click()
On Error GoTo t
'we want to send the contents of txtSend textbox

sock1.SendData txtSend  'trasmits the string to host


'we have send the data to the server by we
'also need to add them to our Chat Buffer
'so we can se what we wrote
txtLog = txtLog & "Client : " & txtSend & vbCrLf

'and then we clear the txtSend textbox so the
'user can write the next message
txtSend = ""

'error handling
'( for example , we will get an error if try to send
'  any data without being connected )
Exit Sub
t:
txtLog.Text = "User Is Not ONLINE" & vbCrLf & "Message Will Not Be Delivered!"
sock1_Close   'close the connection
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Form_Load()
On Error GoTo t

'sock1 is the name of our Winsock ActiveX Control

sock1.Close 'we close it in case it was trying to connect

'txtIP is the textbox holding the host IP
'txtIP can contain both hostnames ( like www.google.com ) or IPs ( like 127.0.0.1 )
sock1.RemoteHost = txtIP    'set the remote host to the ip we wrote
                            'in the txtIP textbox

'txtPort is the textbox holding the Port number
sock1.RemotePort = txtPort  'set the port we want to connect to
                            '( the server must be listening on this port too)
                            
                            
sock1.Connect               'try to connect


Exit Sub
t:
txtLog.Text = "User Is Not ONLINE" & vbCrLf & "Message Will Not Be Delivered!"
End Sub
Public Sub getUser(UName As String)
lblTo.Caption = "Chatting With " & UName
Me.Caption = "Chatting With " & UName
Call useDB(UName)
End Sub

Private Sub sock1_Close()
'handles the closing of the connection

sock1.Close  'close connection

txtLog = txtLog & "*** Disconnected" & vbCrLf

End Sub

Private Sub sock1_Connect()
'txtLog is the textbox used as our
'chat buffer.

'sock1.RemoteHost returns the hostname( or ip ) of the host
'sock1.RemoteHostIP returns the IP of the host

txtLog = "Connected to " & sock1.RemoteHostIP & vbCrLf

End Sub

Private Sub sock1_DataArrival(ByVal bytesTotal As Long)
'This is being trigger every time new data arrive
'we use the GetData function which returns the data that winsock is holding

Dim dat As String     'where to put the data

sock1.GetData dat, vbString   'writes the new data in our string dat ( string format )

'add the new message to our chat buffer
txtLog = txtLog & "Server : " & dat & vbCrLf

End Sub

Private Sub sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'this event is to handle any kind of errors
'happend while using winsock

'Number gives you the number code of that specific error
'Description gives you string with a simple explanation about the error

'append the error message in the chat buffer
txtLog = txtLog & "*** Error : " & Description & vbCrLf

'and now we need to close the connection
sock1_Close

'you could also use sock1.close function but I
'prefer to call it within the Sock1_Close functions that
'handles the connection closing in general

End Sub

