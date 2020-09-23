VERSION 5.00
Object = "*\A..\TRANSL~2\Project1.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   960
   ClientTop       =   1050
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin Project1.Translucency Translucency1 
      Left            =   1800
      Top             =   1110
      _ExtentX        =   847
      _ExtentY        =   847
      BlendColor      =   16761024
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Praveen"
      Height          =   465
      Left            =   3630
      TabIndex        =   0
      Top             =   2610
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   3105
      Index           =   1
      Left            =   60
      Top             =   60
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3165
      Index           =   0
      Left            =   30
      Top             =   30
      Width           =   4635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Translucency1.drawTranslucency
End Sub

