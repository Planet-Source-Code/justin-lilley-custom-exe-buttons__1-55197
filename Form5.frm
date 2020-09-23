VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Msg Box's"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Make"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Text            =   "Please Vote NOW!"
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Text            =   "Did You Vote?"
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Custom Msg Box's!"
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
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " VB Symbol:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MSG box Info:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MSG box title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Frmmain.Show
End Sub

Private Sub Command2_Click()
If Combo1.Text = "vbExclamation" Then ' If you choose vbExclamation
push = MsgBox(Text2.Text, vbExclamation, Text1.Text) ' put it in
End If

If Combo1.Text = "vbCritical" Then '[same]
push = MsgBox(Text2.Text, vbCritical, Text1.Text)
End If

If Combo1.Text = "vbInformation" Then '[same]
push = MsgBox(Text2.Text, vbInformation, Text1.Text)
End If

If Combo1.Text = "vbQuestion" Then '[same]
push = MsgBox(Text2.Text, vbQuestion, Text1.Text)
End If

End Sub

Private Sub Form_Load()
Combo1.AddItem "vbCritical" 'put items in...
Combo1.AddItem "vbExclamation"
Combo1.AddItem "vbInformation"
Combo1.AddItem "vbQuestion"
End Sub

