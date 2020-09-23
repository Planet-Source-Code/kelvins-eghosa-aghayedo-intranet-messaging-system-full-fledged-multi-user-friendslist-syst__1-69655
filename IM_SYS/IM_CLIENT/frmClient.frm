VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Intranet Messaging System - Osazuwa Dumnodi Henrietta [SCN0308563]"
   ClientHeight    =   4125
   ClientLeft      =   2325
   ClientTop       =   2850
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleMode       =   0  'User
   ScaleWidth      =   461.034
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Text            =   "John Accounting"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtIPAddress 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "localhost"
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Text            =   "1212"
      Top             =   3600
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sktClient 
      Left            =   360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtOutput 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "My Username:"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "Port:"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Server IP Address:"
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
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "Connect To Secured IM Firewall Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   2880
      Picture         =   "frmClient.frx":0000
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Server Response:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myName As String
Private Sub AddToOutput(myText As String)
        txtOutput.Text = myText & vbNewLine & txtOutput.Text
End Sub
Private Sub cmdClient_Click()
    sktClient.Close
    Call getInfo
    End Sub
Public Sub getMyname(myName As String)
txtInput.Text = myName
End Sub
Private Sub cmdClose_Click()
    ' close the connection
    If sktClient.State <> sckClosed Then sktClient.Close
    Unload Me
End Sub

Private Sub cmdConnect_Click()

    ' First close the socket
    sktClient.Close
    
    ' Tell the user what's going on
    Call AddToOutput(vbNewLine & "Connecting on " & _
        gIPAddress & ", port " & gPort & vbNewLine)

    ' Connect to the server
    sktClient.Connect
    
    ' Reset my handle
    myName = txtInput.Text
End Sub

Private Sub cmdSend_Click()
    ' Set the name of the client
   ' If myName = "" Then
        
        ' we know the user's been asked to enter their
        ' handle by the server, they'll type it into txtInput
        myName = txtInput.Text
        
        ' Reset my title bar
        Me.Caption = "Intranet Messaging System - Osazuwa Dumnodi Henrietta [SCN0308563] - " & myName
        
     
    ' Send the data
   sktClient.SendData txtInput.Text
    
    ' Zero the input
    'txtInput.Text = ""
 cmdSend.Caption = "Connected To Firewall Server!"
    txtInput.Move (10)
    Label5.Move (10)
    cmdSend.Width = 200
    cmdSend.Move (150)
    cmdSend.Enabled = False
    frmUsers.myName (txtInput.Text)
    Me.Hide
    frmUsers.Show (vbModal)
    
    
End Sub
Private Sub getInfo()
  ' Initiate the connection to the server
  
    gPort = txtPort.Text
    gIPAddress = txtIPAddress.Text
    
    sktClient.RemoteHost = gIPAddress
    sktClient.RemotePort = gPort
    
    Call AddToOutput(vbNewLine & "Connecting on " & gIPAddress & ", port " & gPort & vbNewLine)
    
    ' Connect to the server
        sktClient.Connect
    ' Reset my handle
    myName = txtInput.Text
End Sub

Private Sub Form_Load()
cmdSend.Visible = False
End Sub

Private Sub sktClient_DataArrival(ByVal bytesTotal As Long)
    Dim newData As String
    ' Get the arriving data and print it out.
    sktClient.GetData newData
    ' add the data to the output
    txtOutput.Text = newData & txtOutput.Text
    Dim getSIP As String
    getSIP = gIPAddress
    cmdClient.Visible = False
    cmdSend.Visible = True
    txtIPAddress.Visible = False
    txtPort.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    
    'Me.Hide
    'frmUsers.Show
End Sub
Private Sub sktClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' This should handle any winsock errors.
    Call AddToOutput("Error: " & Description)
End Sub
