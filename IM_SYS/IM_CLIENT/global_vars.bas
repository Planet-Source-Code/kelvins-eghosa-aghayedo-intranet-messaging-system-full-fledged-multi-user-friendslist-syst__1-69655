Attribute VB_Name = "global_vars"
Public gPort As Integer
Public gIPAddress As String
Public gTotalincoming As String
Public gServerLog As String
Public gNumConnections As Long



Public Sub AddToTotalIncoming(additional_text As String)
    
     gTotalincoming = additional_text & vbNewLine & gTotalincoming
    
End Sub
Public Sub AddToServerLog(additional_text As String)
    
     gServerLog = additional_text & vbNewLine & gServerLog
    
End Sub




