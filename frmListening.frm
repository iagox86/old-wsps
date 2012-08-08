VERSION 5.00
Begin VB.Form frmListening 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listening..."
   ClientHeight    =   1185
   ClientLeft      =   6165
   ClientTop       =   5310
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.Timer timer 
      Interval        =   150
      Left            =   120
      Top             =   120
   End
   Begin VB.Label txtPort 
      Caption         =   "0..."
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Listening on port"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label txtimage 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmListening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMain.ws.Close
    Me.Hide
End Sub


Private Sub timer_Timer()
    Static state As Integer
    
    If state = 0 Then
        state = 1
        txtimage.Caption = "--"
    ElseIf state = 1 Then
        state = 2
        txtimage.Caption = "/"
    ElseIf state = 2 Then
        state = 3
        txtimage.Caption = "|"
    Else
        state = 0
        txtimage.Caption = "\"
    End If
    

End Sub
