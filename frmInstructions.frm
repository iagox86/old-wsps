VERSION 5.00
Begin VB.Form frmInstructions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instructions"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmInstructions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmInstructions.frx":030A
      Height          =   735
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "DON'T try to send a packet unless you are connected."
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmInstructions.frx":03C4
      Height          =   975
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmInstructions.frx":04AB
      Height          =   1095
      Index           =   7
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Things you gotta know:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
Me.Hide
End Sub

'Private Sub Form_Load()
'Me.Hide
'End Sub

Private Sub Label1_Click(index As Integer)
Me.Hide
End Sub

