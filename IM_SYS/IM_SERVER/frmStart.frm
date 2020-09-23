VERSION 5.00
Begin VB.Form frmStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "IM_SERVER CONNECTION "
   ClientHeight    =   1290
   ClientLeft      =   6675
   ClientTop       =   3990
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Text            =   "1212"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtIPAddress 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "localhost"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdServer 
      Appearance      =   0  'Flat
      Caption         =   "&Connect"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Port"
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
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Server IP Address"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetAddress()
    gPort = txtPort.Text
    gIPAddress = txtIPAddress.Text
End Sub

Private Sub cmdClient_Click()
    Call SetAddress
  '  frmClient.Show
    Unload Me
End Sub

Private Sub cmdServer_Click()
    Call SetAddress
    frmServer.Show
    Unload Me
End Sub

