VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "1945"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtRemoteHost 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Remote host:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Remote port:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Local port:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RemoteHost As String
Dim RemotePort As String
Dim LocalPort As String
Dim DoFill As Boolean

Private Sub cmdCancel_Click()
    txtRemoteHost.SetFocus
    txtRemoteHost.Text = RemoteHost
    txtRemotePort.Text = RemotePort
    txtLocalPort.Text = LocalPort
    
    Me.Hide
End Sub

Private Sub cmdConnect_Click()
    If (Not (IsNumeric(txtLocalPort))) Then
        MsgBox "Local port must be a numeric value.", vbCritical
    ElseIf (Not (IsNumeric(txtRemotePort.Text))) Then
        MsgBox "Remote port must be a numeric value.", vbCritical
    ElseIf (txtRemoteHost.Text = "") Then
        MsgBox "You must provide a remote host."
    Else
        If frmMain.ws.state = 0 Then
            frmMain.ws.Close
            frmMain.ws.LocalPort = txtLocalPort.Text
            frmMain.ws.RemotePort = Val(txtRemotePort.Text)
            frmMain.ws.RemoteHost = txtRemoteHost.Text
            frmMain.ws.Connect
        Else
            MsgBox "Winsock is already connected.", vbCritical, "Oops!"
        End If
      
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    txtRemoteHost.SetFocus
    RemoteHost = txtRemoteHost.Text
    RemotePort = txtRemotePort.Text
    LocalPort = txtLocalPort.Text
    
End Sub

Private Sub txtRemoteHost_Change()
    'Autofills the remotehost box


    Dim Selection As Integer
    Dim Hist() As String
    Dim iindex As Integer
    
    On Error Resume Next
    
    If DoFill Then
        If Len(txtRemoteHost.Text) > 0 Then
            For iindex = 1 To gHistoryNum
                Hist = Split(gHistoryList(iindex), ":", 2)
                If txtRemoteHost.Text = Left(Hist(0), Len(txtRemoteHost.Text)) Then
                    txtRemotePort.Text = Hist(1)
                    Selection = txtRemoteHost.SelStart
                    txtRemoteHost.Text = Hist(0)
                    txtRemoteHost.SelStart = Selection
                    txtRemoteHost.SelLength = Len(txtRemoteHost) - txtRemoteHost.SelStart
                End If
            Next
        End If
    End If
    
    DoFill = False
End Sub

Private Sub txtRemoteHost_GotFocus()
    txtRemoteHost.SelStart = 0
    txtRemoteHost.SelLength = Len(txtRemoteHost.Text)
End Sub

Private Sub txtRemoteHost_KeyPress(KeyAscii As Integer)
    'if it's a control character, don't do the autofill.. otherwise, put the flag up

    If KeyAscii > &H20 Then
        DoFill = True
    Else
        DoFill = False
    End If
End Sub

Private Sub txtRemotePort_GotFocus()
    txtRemotePort.SelStart = 0
    txtRemotePort.SelLength = Len(txtRemotePort.Text)
End Sub
Private Sub txtLocalPort_GotFocus()
    txtLocalPort.SelStart = 0
    txtLocalPort.SelLength = Len(txtLocalPort.Text)
End Sub

