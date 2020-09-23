VERSION 5.00
Begin VB.Form frmUMan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Intranet Friends Management"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "My IM Friends"
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command21 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&New"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Next"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Prev"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox t3 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox t2 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox t1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lb 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUMan"
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
    
    rs.MoveFirst
    t1.Text = rs(0)
    t2.Text = rs(1)
    t3.Text = rs(2)
    
    'rs.Close
    End Sub

Private Sub cmdsave_Click()
    rs.Update
    rs(0) = t1.Text
    rs(1) = t2.Text
    rs(2) = t2.Text
    lb.Caption = "Saved!"
End Sub

Private Sub Command1_Click()
lb.Caption = ""
    rs.MovePrevious
    If rs.BOF = True Then
     MsgBox "This is the last record.", vbExclamation, "Note it..."
     rs.MoveFirst
    End If
    t1.Text = rs(0)
    t2.Text = rs(1)
    t3.Text = rs(2)
End Sub

Private Sub Command21_Click()
        If (t1.Text = "" Or t2.Text = "" Or t3.Text = "") Then
        lb.Caption = "Error!"
        Else
        rs.AddNew
        rs(0) = t1.Text
        rs(1) = t2.Text
        rs(2) = t2.Text
        lb.Caption = "Added!"
        End If
End Sub

Private Sub Command3_Click()
lb.Caption = ""
 rs.MoveNext
    If rs.EOF = True Then
        MsgBox "This is the last record.", vbExclamation, "Note it..."
        rs.MoveLast
    End If
t1.Text = rs(0)
    t2.Text = rs(1)
    t3.Text = rs(2)
End Sub

Private Sub Command4_Click()
t1.Text = ""
t2.Text = ""
t3.Text = ""
cmdsave.Caption = "Add"
End Sub

Private Sub Form_Load()
Call useDB
Command21.Visible = False
End Sub
