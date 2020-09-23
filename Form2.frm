VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Up"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":030A
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Button3"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Button2"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Button1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Done"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "AIM"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin MSComDlg.CommonDialog setup 
      Left            =   4080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "C:\Path"
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button Name:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Path:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
setup.ShowOpen ' opens dialog
Text1.Text = setup.FileName ' puts path in text file
End Sub

Private Sub Command2_Click()
Unload Me ' unloads settings
Form1.Show
End Sub

Private Sub Command3_Click()
If Option1.Value = True Then ' if option1 is choose
WriteIniData "Button1", "Button1", Text2.Text, App.Path & "/button.ini" ' [see middle]
WriteIniData "Button1", "Path", Text1.Text, App.Path & "/button.ini" 'Write in exe path
End If
If Option2.Value = True Then ' if option2 is choose
         'Place under,  Under,  write text2 in the ini file, by app in button.ini
WriteIniData "Button2", "Button2", Text2.Text, App.Path & "/button.ini"
WriteIniData "Button2", "Path", Text1.Text, App.Path & "/button.ini" 'Write in exe path
End If
If Option3.Value = True Then ' if option3 is choose
WriteIniData "Button3", "Button3", Text2.Text, App.Path & "/button.ini" ' [see middle]
WriteIniData "Button3", "Path", Text1.Text, App.Path & "/button.ini" 'Write in exe path
End If
End Sub

Private Sub Form_Load()
'[ all below writes in the data of pre saved settings!]
Option1.Value = True

If Option1.Value = True Then
Text2.Text = GetIniData("Button1", "button1", "", App.Path & "/Button.ini")
Text1.Text = GetIniData("Button1", "path", "", App.Path & "/Button.ini")
End If
    If Option2.Value = True Then
    Text2.Text = GetIniData("Button2", "button2", "", App.Path & "/Button.ini")
    Text1.Text = GetIniData("Button2", "path", "", App.Path & "/Button.ini")
    End If
If Option3.Value = True Then
Text2.Text = GetIniData("Button3", "button3", "", App.Path & "/Button.ini")
Text1.Text = GetIniData("Button3", "path", "", App.Path & "/Button.ini")
End If
End Sub

Private Sub Timer1_Timer()
Form2.Refresh ' Refresh the form
End Sub
