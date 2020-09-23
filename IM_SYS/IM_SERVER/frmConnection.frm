VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection"
   ClientHeight    =   615
   ClientLeft      =   7635
   ClientTop       =   2790
   ClientWidth     =   1215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   1215
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer timMyTimer 
      Interval        =   200
      Left            =   720
      Top             =   120
   End
   Begin MSWinsockLib.Winsock sktListener 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myTotalIncomingCopy As String
Dim myID As Long
Dim myhandle As String

Private Sub Form_Load()
    myhandle = "unknown"
    sktListener.LocalPort = gPort
    sktListener.Listen
    AddToServerLog ("New Connection Form is now listening")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'if the connection is not closed then close it.
    If sktListener.State <> 0 Then sktListener.Close
End Sub

Private Sub sktListener_ConnectionRequest(ByVal requestID As Long)

    ' If the connection is not closed then close it
    If sktListener.State <> 0 Then sktListener.Close
    
    ' Now accept the new connection
    sktListener.LocalPort = 0
    sktListener.Accept requestID
    
    myID = requestID
    AddToServerLog ("Connection: Made connection reqID = " & myID)
    
    sktListener.SendData "Connection made, what's your handle?"
    myhandle = "unknown"
    
    ' Tell the world there's a new connection needed
    Call StartNewConnection
    
End Sub

Private Sub sktListener_DataArrival(ByVal bytesTotal As Long)

   Dim newData As String
    
    sktListener.GetData newData
    
    'if we have no handle then they'll have typed it in
    If myhandle = "unknown" Then
        myhandle = newData
        sktListener.SendData vbNewLine & "Welcome to CornflakeChat " & myhandle & vbNewLine
    Else
        AddToTotalIncoming (myhandle & " : " & newData)
    End If
 
End Sub

Private Sub sktListener_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call AddToServerLog("Error on Connection : " & myID & " - " & Description)
End Sub

Private Sub timMyTimer_Timer()

    Dim myDatatoSend As String

    'if there's a skt open then send the other peoples' messages
    If sktListener.State = sckConnected Then
        'open
        
        'only send the stuff I haven't seen
        myDatatoSend = Left(gTotalincoming, Len(gTotalincoming) - Len(myTotalIncomingCopy))
        myTotalIncomingCopy = gTotalincoming
    
        sktListener.SendData myDatatoSend
    End If
End Sub

