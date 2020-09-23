VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Secured IM Client Users"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleLeft       =   1000
   ScaleMode       =   0  'User
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Caption         =   "Registered Users"
      Height          =   7695
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Disconnect / LogOut"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   6360
         Width           =   2055
      End
      Begin MSComctlLib.ListView lstv 
         Height          =   5055
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   8916
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483625
         BackColor       =   -2147483629
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "My Friends List"
            Object.Width           =   4480
         EndProperty
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   2160
         Picture         =   "frmUsers.frx":0000
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "Click To Manage Friends"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Index           =   2
         Left            =   720
         MousePointer    =   2  'Cross
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000014&
         Caption         =   "[SCN0308563]"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   7320
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "This application was developed by:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   6840
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
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
         Left            =   240
         TabIndex        =   4
         Top             =   7080
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "Double-Click on a User To Start Chat Session:"
         ForeColor       =   &H000080FF&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "Double-Click on a User To Start Chat Session:"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmUsers"
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
    str = "select * from users"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    
    While Not rs.EOF()
        lstv.ListItems.Add = rs(0)
        rs.MoveNext
    Wend
    
    rs.Close
    End Sub
    Public Sub myName(getName As String)
    lblName.Caption = "[ " & getName & " ]"
        End Sub
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call useDB
End Sub

Private Sub Label1_Click(Index As Integer)
frmUMan.Show (vbModal)
End Sub

Private Sub lstv_DblClick()
Dim myName As String
myName = lstv.SelectedItem.Text
frmChat.getUser (myName)
frmChat.Show (vbModal)
End Sub
