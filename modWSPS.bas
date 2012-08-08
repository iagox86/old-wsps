Attribute VB_Name = "modWSPS"
Option Explicit

Public gConnected As Boolean
Public gSync As Boolean
Public gLongLines As Boolean
Public Reg
Public gHistory As Boolean
Public gHistoryNum As Integer
Public gHistoryList() As String
Public gUDP As Boolean
Public gAscii(0 To &HFF) As String * 1
Public gDefaultAscii(0 To &HFF) As String * 1

Sub Main()
    Dim iindex As Integer
    'Set up the registry class thing:
    Set Reg = CreateObject("wscript.shell")
    
    'Set the default values for the globals:
    gConnected = False
    gSync = True
    gLongLines = False
    gHistory = True
    gHistoryNum = 4
    
    'If the registry values haven't been done yet,
    'don't crash!
    On Error Resume Next
    'Read the values from the registry, if they exist:
    gSync = Reg.regread("HKEY_CURRENT_USER\Software\WSPS\Ron\gSync")
    If (Not (gSync)) Then
        gLongLines = Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\gLongLines")
    End If
    gHistory = Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\gHistory")
    If (Not (gHistory)) Then
        gHistoryNum = 0
    Else
        gHistoryNum = Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\gHistoryNum")
        If gHistoryNum > 10 Then gHistoryNum = 10
        If gHistoryNum < 0 Then gHistoryNum = 0
        
        ReDim gHistoryList(1 To gHistoryNum)
        
        For iindex = 1 To gHistoryNum
            gHistoryList(iindex) = Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex)
        Next
        
    End If
    
    gUDP = Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\gUDP")
    
    If gUDP = True Then
        frmMain.ws.Protocol = sckUDPProtocol
    Else
        frmMain.ws.Protocol = sckTCPProtocol
    End If
        
    frmMain.Show
    

End Sub


Function setchk(value As Boolean)
    If value = True Then
        setchk = 1
    Else
        setchk = 0
    End If
End Function

Function setbol(value As Integer)
    If value = 1 Then
        setbol = True
    Else
        setbol = False
    End If
End Function
Sub AddToHistory()
    'Adds to the beginning of gHistoryList, moving the rest of the elements up one,
    'but if it's already in the history then just move it to the front
    Dim iindex As Integer
    For iindex = gHistoryNum - 1 To 1 Step -1
        gHistoryList(iindex + 1) = gHistoryList(iindex)
    Next
    
    gHistoryList(1) = frmConnect.txtRemoteHost.Text & ":" & frmConnect.txtRemotePort.Text
    
    SaveHistory
    
    
End Sub

Sub SaveHistory()
    'saves the history to the registry
    Dim iindex As Integer
    
    On Error Resume Next
    
    For iindex = 1 To 10
        If iindex <= gHistoryNum Then
            Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex, gHistoryList(iindex)
        Else
            Reg.regdelete "HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex
        End If
    Next

End Sub
