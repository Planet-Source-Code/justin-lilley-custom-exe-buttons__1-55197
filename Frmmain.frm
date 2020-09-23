VERSION 5.00
Begin VB.Form Frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome        !~`[Justin]`~!"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Frmmain.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "G&o Here"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go Here"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ahoy to all please enjoy!"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Msg Boxs!"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Buttons that Run EXE's"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      X1              =   0
      X2              =   4680
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Frmmain.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
Frmmain.Hide
Form5.Show
End Sub

Private Sub Command3_Click()
Unload Me   '[Ends It all]
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End
End Sub
