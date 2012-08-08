VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3975
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   4095
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   0
      Left            =   210
      ScaleHeight     =   2700
      ScaleWidth      =   3675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   3675
      Begin VB.Frame fraSample1 
         Caption         =   "Synchronization"
         Height          =   2625
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   3615
         Begin VB.CheckBox chkSync 
            Caption         =   "Synchronize recieved data"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            ToolTipText     =   $"frmOptions.frx":000C
            Top             =   360
            Value           =   1  'Checked
            Width           =   2655
         End
         Begin VB.CheckBox chkLongLines 
            Caption         =   "Allow long lines on ascii"
            Enabled         =   0   'False
            Height          =   255
            Left            =   480
            TabIndex        =   9
            ToolTipText     =   "The 16-character limitation is removed from the recieved ascii box"
            Top             =   840
            Width           =   2655
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   3495
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3495
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   450
      TabIndex        =   1
      Top             =   3495
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   1
      Left            =   210
      ScaleHeight     =   2700
      ScaleWidth      =   3675
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   600
      Width           =   3675
      Begin VB.Frame fraSample2 
         Caption         =   "History"
         Height          =   2625
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   3615
         Begin VB.TextBox txtHistory 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   13
            Text            =   "4"
            Top             =   1065
            Width           =   375
         End
         Begin VB.CheckBox chkHistory 
            Caption         =   "Remember history?"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.Label lblHistory2 
            Caption         =   "(1 - 10)"
            Height          =   255
            Left            =   720
            TabIndex        =   12
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblHistory1 
            Caption         =   "Number of connections to remember?"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   2775
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2700
      Index           =   2
      Left            =   210
      ScaleHeight     =   2700
      ScaleWidth      =   3675
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   3675
      Begin VB.Frame Frame1 
         Caption         =   "TCP or UDP"
         Height          =   2625
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3615
         Begin VB.OptionButton optUDP 
            Caption         =   "User Datagram Protocol"
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   2655
         End
         Begin VB.OptionButton optTCP 
            Caption         =   "Transfer Control Protocol"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5794
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Display"
            Key             =   "Display"
            Object.ToolTipText     =   "Set display options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            Key             =   "History"
            Object.ToolTipText     =   "History options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Connection"
            Key             =   "Connection"
            Object.ToolTipText     =   "TCP or UDP"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkHistory_Click()
    If chkHistory.value = 0 Then
        txtHistory.Enabled = False
        txtHistory.Text = "0"
        lblHistory1.Enabled = False
        lblHistory2.Enabled = False
    Else
        txtHistory.Enabled = True
        lblHistory1.Enabled = True
        lblHistory2.Enabled = True
        txtHistory.SelStart = 0
        txtHistory.SelLength = Len(txtHistory.Text)
    End If
End Sub

Private Sub chkSync_Click()
    chkLongLines.Enabled = Not setbol(chkSync.value)
    
    If (chkSync.value) Then
        chkLongLines.value = 0
    End If
    
    
    'MsgBox setbol(chkSync) & " " & setchk(chkLongLines.Enabled)
End Sub

Private Sub cmdApply_Click()
    Dim iindex As Integer

    
    If (gSync <> setbol(chkSync.value)) Then
        If (MsgBox("Changing the synchronization settings will reset recieved packet boxes." & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbExclamation, "Warning...") = vbYes) Then
            frmMain.lstAscii.Text = ""
            frmMain.lstHex.Text = ""
            frmMain.lstAscii.Tag = 0
            frmMain.lstHex.Tag = 0
            gSync = setbol(chkSync.value)
        End If
    End If
        
    gUDP = optUDP.value
    gLongLines = setbol(chkLongLines.value)
    gHistory = setbol(chkHistory.value)
    gHistoryNum = txtHistory.Text
    If gHistoryNum > 0 Then
        ReDim Preserve gHistoryList(1 To gHistoryNum) As String
    End If
    SaveHistory
    If gUDP = True Then
        frmMain.ws.Protocol = sckUDPProtocol
    Else
        frmMain.ws.Protocol = sckTCPProtocol
    End If
    
    'Write the new values to the registry
    Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\gSync", setchk(gSync)
    Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\gLongLines", setchk(gLongLines)
    Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\gHistory", setchk(gHistory)
    Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\gHistoryNum", gHistoryNum
    Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\gUDP", gUDP
    
    'Get rid of unneeded history keys or add new ones
    On Error Resume Next
    
    For iindex = 1 To gHistoryNum
        If Reg.regread("HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex) = "" Then
            Reg.regwrite "HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex, ""
        End If
    Next
    For iindex = gHistoryNum + 1 To 10
        Reg.regdelete ("HKEY_CURRENT_USER\Software\Ron\WSPS\History\history" & iindex)
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call cmdApply_Click

    Unload Me
End Sub

Private Sub Form_Activate()
    
    optTCP.value = Not (gUDP)
    optUDP.value = gUDP
    
    chkSync.value = setchk(gSync)
    chkLongLines.value = setchk(gLongLines)
    chkLongLines.Enabled = Not setbol(chkSync.value)
    chkHistory.value = setchk(gHistory)
    txtHistory.Text = gHistoryNum
    Call chkHistory_Click
    
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub

Private Sub txtHistory_Click()
    txtHistory.SelStart = 0
    txtHistory.SelLength = Len(txtHistory.Text)
End Sub

Private Sub txtHistory_KeyPress(KeyAscii As Integer)

    If Not ((KeyAscii >= &H30 And KeyAscii <= &H39) Or KeyAscii = &H8) Then
        KeyAscii = 0
    End If
    
End Sub
