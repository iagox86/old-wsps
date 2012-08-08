VERSION 5.00
Begin VB.Form frmListen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listen"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Listen"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtLocalPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "1945"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Local port:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LocalPort As String

Private Sub cmdCancel_Click()
    txtLocalPort.Text = LocalPort
    Me.Tag = "0"
    Me.Hide
End Sub

Private Sub cmdlisten_Click()
    If (Not (IsNumeric(txtLocalPort))) Then
        MsgBox "Local port must be a numeric value.", vbCritical
    Else
        If (frmMain.ws.state = 0) Then
            frmMain.ws.Close
            frmMain.ws.LocalPort = txtLocalPort.Text
            frmMain.ws.Listen
            frmListening.txtPort.Caption = txtLocalPort.Text & "..."
            Me.Tag = "1"
            Reg.regwrite "HKEY_CURRENT_USER\Software\ron\WSPS\listenport", txtLocalPort.Text
        Else
            MsgBox "Winsock is already connected.", vbCritical
        End If




        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next

    Me.Tag = "0"
    
    txtLocalPort.Text = Reg.regread("HKEY_CURRENT_USER\software\ron\WSPS\listenport")
    
    txtLocalPort.SetFocus
    LocalPort = txtLocalPort.Text
    
End Sub
Private Sub txtLocalPort_GotFocus()
    txtLocalPort.SelStart = 0
    txtLocalPort.SelLength = Len(txtLocalPort.Text)
End Sub
