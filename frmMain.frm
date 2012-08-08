VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winsock packet sender Beta"
   ClientHeight    =   7605
   ClientLeft      =   2520
   ClientTop       =   2040
   ClientWidth     =   9210
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   9210
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lstAscii 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   6720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox lstHex 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   360
      Width           =   6375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear recieved packets"
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdBinary 
      Caption         =   "Send Binary Packet"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   5040
      Width           =   6375
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Ascii Packet"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtAscii 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   6720
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblSAscii 
      Caption         =   "Ascii Packet:"
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblSHex 
      Caption         =   "Hex Packet:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblRHex 
      Caption         =   "Recieved packets, hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblRAscii 
      Caption         =   "Recieved packets, ascii:"
      Height          =   255
      Left            =   6720
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   10
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuComment 
         Caption         =   "&Send Question/Comment/Bug report..."
      End
      Begin VB.Menu mnuOptionsPreferences 
         Caption         =   "&Preferences..."
      End
   End
   Begin VB.Menu mnuConnection 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnectionConnect 
         Caption         =   "&Connect..."
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuConnectionListen 
         Caption         =   "&Listen..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuConnectionDisconnect 
         Caption         =   "&Disconnect"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInstructions 
         Caption         =   "&Instructions"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Synchronization vs. long lines
    'When synchronization is enabled, the hex and the ascii stay perfectly
    'matched in the textboxes, including which line they go on, and the
    'crlf's from the string become "\n"'s
    
    'When synchronization is disabled, the crlfs are left alone, but the
    'lines are still a maximum of 16 characters lone
    
    'When Long lines are enabled, the 16 character limit is removed and
    'the lines continue until a crlf is found.  Synchronization must be
    'disabled for this to work, otherwise it would just be one long line

'New idea!

'I'm going to have 3 different arrays, and each will hold one thing (obviously)
'one is for the hex, one is for the ascii, and one is for decimal.
'When data is recieved it'll be put into the Decimal array and from there
'it'll be converted to hex and ascii and put into the respective boxes.

'This will be useful for synchronizing the boxes.
Option Explicit
Dim mDecimal() As Integer
Dim mHex() As String
Dim mAscii() As String
Dim sAscii As String
Dim sHex As String

Dim FormWidth As Integer
Dim FormHeight As Integer


'Each row in the hex box starts at a multiple of 50:
Const mHexLength As Integer = 50
'Each row in the ascii box starts at a multiple of 18:
Const mAsciiLength As Integer = 18

Private Sub cmdBinary_Click()
    Dim iindex As Integer
    Dim SendString As String
    Dim Buffer() As String
    Dim temp As Integer
    Dim HexText As String
    
    On Error Resume Next
    
    txtHex.Text = UCase(txtHex.Text)
    HexText = Replace(txtHex.Text, vbCrLf, " ")
    
    Buffer = Split(HexText, " ")
    SendString = ""
    
    For iindex = LBound(Buffer) To UBound(Buffer)
        If (Buffer(iindex) <> "") Then
            If (Len(Buffer(iindex)) <= 2) Then
        
                temp = 0
                
                If (IsNumeric(Left(Buffer(iindex), 1))) Then
                    temp = Val(Left(Buffer(iindex), 1))
                ElseIf Left(Buffer(iindex), 1) = "A" Then
                    temp = 10
                ElseIf Left(Buffer(iindex), 1) = "B" Then
                    temp = 11
                ElseIf Left(Buffer(iindex), 1) = "C" Then
                    temp = 12
                ElseIf Left(Buffer(iindex), 1) = "D" Then
                    temp = 13
                ElseIf Left(Buffer(iindex), 1) = "E" Then
                    temp = 14
                ElseIf Left(Buffer(iindex), 1) = "F" Then
                    temp = 15
                Else
                    MsgBox "An invalid character was found.  The character is " & Left(Buffer(iindex), 1) & " and has the code " & Hex(Asc(Left(Buffer(iindex), 1))), vbCritical
                End If
                
                If (Len(Buffer(iindex)) = 2) Then
                    temp = temp * 16
                    
                    If (IsNumeric(Right(Buffer(iindex), 1))) Then
                        temp = temp + Val(Right(Buffer(iindex), 1))
                    ElseIf Right(Buffer(iindex), 1) = "A" Then
                        temp = temp + 10
                    ElseIf Right(Buffer(iindex), 1) = "B" Then
                        temp = temp + 11
                    ElseIf Right(Buffer(iindex), 1) = "C" Then
                        temp = temp + 12
                    ElseIf Right(Buffer(iindex), 1) = "D" Then
                        temp = temp + 13
                    ElseIf Right(Buffer(iindex), 1) = "E" Then
                        temp = temp + 14
                    ElseIf Right(Buffer(iindex), 1) = "F" Then
                        temp = temp + 15
                    Else
                        MsgBox "An invalid character was found.  The character is " & Right(Buffer(iindex), 1) & " and has the code " & Hex(Asc(Right(Buffer(iindex), 1))), vbCritical
                    End If
                End If
                
                SendString = SendString & Chr(temp)
            Else
                MsgBox Buffer(iindex) & " is invalid; all hex values must have 1 or 2 values.  This will be skipped.", vbCritical, "Error"
            End If
        End If
    Next

    ws.SendData SendString
   
    txtHex.SelStart = 0
    txtHex.SelLength = Len(txtHex.Text)
    
    txtHex.SetFocus

End Sub

Private Sub cmdClear_Click()
    lstHex.Text = ""
    lstAscii.Text = ""
    lstHex.Tag = 0
    lstAscii.Tag = 0
End Sub

Private Sub cmdSend_Click()
    On Error Resume Next
    ws.SendData txtAscii.Text
    txtAscii.SelStart = 0
    txtAscii.SelLength = Len(txtAscii.Text)
    txtAscii.SetFocus
End Sub

Private Sub Form_Activate()
    ShowHistory
End Sub



Private Sub Form_Paint()
    frmMain.Width = 9300
    frmMain.Height = 8265

End Sub

Private Sub Form_Terminate()
    ws.Close
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ws.Close
    End
End Sub

Private Sub lstAscii_Change()
    txtHex.TabStop = False
    txtAscii.TabStop = False
    
    lstHex.TabStop = True
    lstAscii.TabStop = True

End Sub

Private Sub lstAscii_Click()
    Call MoveHex
End Sub

Private Sub lstAscii_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveHex
End Sub

Private Sub lstAscii_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub lstAscii_KeyUp(KeyCode As Integer, Shift As Integer)
    Call MoveHex
End Sub

Private Sub lstAscii_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveHex
End Sub

Private Sub lstAscii_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button Then MoveHex
End Sub

Private Sub lstAscii_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveHex
End Sub

Private Sub lstHex_Change()
    txtHex.TabStop = False
    txtAscii.TabStop = False
    
    lstHex.TabStop = True
    lstAscii.TabStop = True
End Sub

Private Sub lstHex_Click()
    Call MoveAscii
End Sub

Private Sub lstHex_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MoveAscii
End Sub

Private Sub lstHex_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub lstHex_KeyUp(KeyCode As Integer, Shift As Integer)

    Call MoveAscii
End Sub

Private Sub lstHex_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveAscii
End Sub

Private Sub lstHex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button Then MoveAscii
End Sub

Private Sub lstHex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MoveAscii
End Sub

Private Sub mnuComment_Click()
    frmComment.Show vbModal
    
End Sub

Private Sub mnuConnectionConnect_Click()
    frmConnect.Show vbModal
End Sub

Private Sub mnuConnectionDisconnect_Click()
    If (ws.state = 0) Then
        MsgBox "Winsock isn't connected", vbCritical
    Else
        ws.Close
        Call ws_Close
    End If
End Sub

Private Sub mnuConnectionListen_Click()
    frmListen.Show vbModal
    
    If frmListen.Tag = "1" Then
        frmListening.Show vbModeless
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    ws.Close
    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpInstructions_Click()
    Load frmInstructions
    frmInstructions.Show vbModal
End Sub

Private Sub mnuOptionsPreferances_Click()
    Load frmOptions
    frmOptions.Show vbModeless
End Sub

Private Sub mnuHistory_Click(index As Integer)
    Dim sHistory() As String
    
    On Error Resume Next
    
    sHistory = Split(mnuHistory(index).Caption, ":", 2)
    
    frmConnect.txtRemoteHost.Text = sHistory(0)
    frmConnect.txtRemotePort.Text = sHistory(1)

    frmConnect.Show vbModal
    
    
End Sub

Private Sub mnuOptionsPreferences_Click()
    frmOptions.Show vbModal
End Sub

Private Sub txtAscii_Change()
    cmdSend.Caption = "&Send Ascii Packet"
    cmdBinary.Caption = "Send Binary Packet"

    txtHex.TabStop = True
    txtAscii.TabStop = True
    
    lstHex.TabStop = False
    lstAscii.TabStop = False
End Sub

Private Sub txtHex_Change()
    cmdBinary.Caption = "&Send Binary Packet"
    cmdSend.Caption = "Send Ascii Packet"
    
    txtHex.TabStop = True
    txtAscii.TabStop = True
    
    lstHex.TabStop = False
    lstAscii.TabStop = False
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
        
    
    If (KeyAscii > &H60 And KeyAscii < &H67) Then
        KeyAscii = KeyAscii - &H20
    End If
    
    If Not ((KeyAscii >= &H30 And KeyAscii <= &H39) Or (KeyAscii > &H40 And KeyAscii < &H47) Or KeyAscii = &H20 Or KeyAscii = &H8 Or KeyAscii = &HD Or KeyAscii = &HA) Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub ws_Close()
    If gConnected Then
        mnuConnectionConnect.Enabled = True
        mnuConnectionListen.Enabled = True
        gConnected = False
        MsgBox "Winsock has been closed", vbInformation
    End If

End Sub

Sub ws_Connect()
    MsgBox "Winsock has connected successfully!", vbInformation
    lstAscii.Enabled = True
    lstHex.Enabled = True
    mnuConnectionConnect.Enabled = False
    mnuConnectionListen.Enabled = False
    gConnected = True
    
    AddToHistory
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    ws.Accept requestID
    frmListening.Hide
    'MsgBox "Connection successfully recieved!", vbInformation, "Success!"
    
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim Buffer As String
    ws.GetData Buffer, vbString, bytesTotal
    
    ProcessData (Buffer)
    
    
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock error!" & vbCrLf & "Error no: " & Number & vbCrLf & "Description: " & Description, vbCritical, "Winsock Error"
    ws.Close
    Call ws_Close
End Sub

Private Sub ProcessData(data As String)
    Dim iindex As Integer
    
    'resizes the arrays
    ReDim Preserve mDecimal(1 To Len(data) + 1) As Integer
    ReDim Preserve mHex(1 To Len(data) + 1) As String
    ReDim Preserve mAscii(1 To Len(data) + 1) As String
    
    'Fills the arrays
    For iindex = 1 To Len(data)
            'Fills the decimal array based on the recieved data
            mDecimal(iindex) = Asc(Mid(data, iindex, 1))
            
            'Hex:
            If (mDecimal(iindex) < 16) Then
                mHex(iindex) = "0" & Hex(mDecimal(iindex))
            Else
                mHex(iindex) = Hex(mDecimal(iindex))
            End If
            
            'Ascii:
            If mDecimal(iindex) >= &H20 And mDecimal(iindex) Then
                mAscii(iindex) = Chr(mDecimal(iindex))
            Else
                mAscii(iindex) = "."
            End If
            
            If (gSync) Then
                If (mAscii(iindex) = vbCr) Then
                    mAscii(iindex) = "."
                ElseIf (mAscii(iindex) = vbLf) Then
                    mAscii(iindex) = "."
                ElseIf (mAscii(iindex) = vbTab) Then
                    mAscii(iindex) = "."
                End If
            End If
    Next
    
    
    'Fill the textboxes based on the recieved ascii and hex data:
    For iindex = LBound(mHex) To UBound(mHex)
        If mHex(iindex) <> "" Then
            If (lstHex.Tag > 15 And mHex(iindex) <> "") Then
                sHex = sHex & vbCrLf
                lstHex.Tag = 1
            ElseIf mHex(iindex) <> "" Then
                lstHex.Tag = lstHex.Tag + 1
            End If
            
            'Add the hex as normal, the "0" in small values is included in the
            'array already, but that might change...
            If mHex(iindex) <> "" Then
                sHex = sHex & mHex(iindex) & " "
            End If
            'If Len(sHex) > 8 Then
            '    Replace sHex, "  ", " ", Len(sHex) - 8
            'End If
            
            'Every 16th ascii value goes on the next value if it's short
            'and synchronized, so:
            If (gLongLines = False) Then
                If (lstAscii.Tag > 15) Then
                    sAscii = sAscii & vbCrLf
                    lstAscii.Tag = 1
                Else
                    lstAscii.Tag = lstAscii.Tag + 1
                End If
            End If
            
            
            
            If iindex > LBound(mAscii) Then
                If (mAscii(iindex) = vbLf And mAscii(iindex - 1) = vbCr) Then
                    lstAscii.Tag = 1
                End If
            End If
                
            
            sAscii = sAscii & mAscii(iindex)
            
            'DoEvents
        End If
    Next
    
    lstHex.Text = lstHex.Text & sHex
    lstAscii.Text = lstAscii.Text & sAscii
    
    sHex = ""
    sAscii = ""
    
End Sub


Private Sub MoveHex()
    Dim Line As Integer, EndLine As Integer
    Dim Char As Integer, EndChar As Integer
    Dim NewPosition As Integer, NewEnd As Integer
    
    On Error Resume Next
    
    If (gSync = True) Then
        'First we have to find out which character and which line we are
        'on in the ascii box.
        'The line = int((position in the box) / (number of characters in each line))
        Line = Int(lstAscii.SelStart / mAsciiLength)
        EndLine = Int((lstAscii.SelStart + lstAscii.SelLength) / mAsciiLength)
        'The character = (position in the box) % (number of characters in each line)
        Char = ModDiv(lstAscii.SelStart, mAsciiLength)
        EndChar = ModDiv((lstAscii.SelStart + lstAscii.SelLength), mAsciiLength)
        
        NewPosition = Line * mHexLength
        NewPosition = NewPosition + (Char * 3)

        NewEnd = EndLine * mHexLength
        NewEnd = NewEnd + ((EndChar) * 3)
        
        lstHex.SelStart = NewPosition
        lstHex.SelLength = (NewEnd - NewPosition)

        
    End If
End Sub
Private Sub MoveAscii()
    Dim Line As Integer, EndLine As Integer
    Dim Char As Integer, EndChar As Integer
    Dim NewPosition As Integer, NewEnd As Integer


    If (gSync = True) Then
        'First we determine the line and character that we are on in the
        'hex box.
        'Line = int((position in the box) / (number of characters in each line))
        Line = Int(lstHex.SelStart / mHexLength)
        EndLine = Int((lstHex.SelStart + lstHex.SelLength) / mHexLength)
        
        'The character = (the position in the box) % (The number of characters in each line)
        Char = ModDiv(lstHex.SelStart, mHexLength)
        EndChar = ModDiv((lstHex.SelStart + lstHex.SelLength), mHexLength)

        'Because each ascii character is represented by 3 hex characters
        '("## "), we have to divide by 3 to get the exact result:
        Char = Int(Char / 3)
        EndChar = Int((EndChar / 3) + 0.9)
        
        NewPosition = Line * mAsciiLength
        NewPosition = NewPosition + Char
        
        NewEnd = EndLine * mAsciiLength
        NewEnd = NewEnd + EndChar
    
        lstAscii.SelStart = NewPosition
        lstAscii.SelLength = (NewEnd - NewPosition)
    End If
End Sub

Function ModDiv(a As Variant, b As Variant)
    If b = 0 Then
        ModDiv = &HBAD
    Else
        ModDiv = ((a / b) - Int(a / b)) * b
    End If
    
End Function


Sub ShowHistory()
    'Shows the history in the File menu
    
    Dim Vis As Boolean
    Dim iindex As Integer
    
    For iindex = 1 To 10
        Vis = (iindex <= gHistoryNum)
        mnuHistory(iindex).Visible = Vis
        If iindex <= gHistoryNum Then
            mnuHistory(iindex).Caption = gHistoryList(iindex)
        End If
    Next
    
    mnuFileSep.Visible = (gHistoryNum > 0)
    
    
End Sub
