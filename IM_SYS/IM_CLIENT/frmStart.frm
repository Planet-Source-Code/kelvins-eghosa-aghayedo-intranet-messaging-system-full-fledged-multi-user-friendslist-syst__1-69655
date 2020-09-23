VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Machine User Authentication"
   ClientHeight    =   1275
   ClientLeft      =   6675
   ClientTop       =   3990
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4980
   Begin VB.TextBox txtP 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtU 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "&Access IM"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sktClient 
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblr 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Password:"
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
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   4200
      Picture         =   "frmStart.frx":0000
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "IM User Authentication"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Username"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "frmStart"
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
Private Sub useDB()
  Dim str As String
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=dbUsers.mdb;Persist Security Info=False"
    str = "select * from users where username='" & txtU.Text & "' AND password = '" & txtP.Text & "'"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    
    If (rs.EOF = True) Then
       lblr.Caption = "Login Failed!"
    Else
        Call SetAddress
        Dim myUser As String
        myUser = txtU.Text
        frmClient.getMyname (myUser)
        frmClient.Show
        Unload Me
    End If
    rs.Close
    End Sub
Private Sub SetAddress()
'    gPort = txtU.Text
   ' gIPAddress = txtP.Text
End Sub

Private Sub cmdClient_Click()
    Call useDB
End Sub

Private Sub cmdServer_Click()
  
End Sub

Private Sub Label2_Click()

End Sub

