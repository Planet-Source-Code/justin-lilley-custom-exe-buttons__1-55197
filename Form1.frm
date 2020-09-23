VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Buttons!"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "&Back to Main Screen"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Set Up"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Cusomizable buttons!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   4215
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhow 
         Caption         =   "How To..."
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'====== Cutom Buttoms that Run EXE files! =====
'======== AND SAVES IT TO .ini FILES ==========
'==============================================
'|\|/| |\|/| |\|/| |\|/| |\|/| |\|/| |\|/| |\|/|
'             Copyright  2004                  |
'      Code by Justin Lilley aka 715           |
'      Graphics also by Justin Lilley          |
'This code was submitted to Planet Source Code |
'        (www.planetsourcecode.com)            |
'If you downloaded this elsewhere, It is Stolen!
'                                              |
'            THIS CODE IS FREEWARE!            |
'                                              |
' If you wish to use this code or any part of  |
'this code for your big App.  Your wish is then|
'               Hereby Granted!                |
'   (Just be sure to put my name into it!)     |
'                 PLEASE VOTE!                 |
Dim push As String


Private Sub Command1_Click()
On Error Resume Next
Command1.Caption = GetIniData("Button1", "Button1", "Button1", App.Path & "/Button.ini") 'Insert Command Caption!
Shell GetIniData("Button1", "path", "", App.Path & "/Button.ini") 'Runs EXE file
End Sub

Private Sub Command2_Click()
Form2.Show 'Show setup
Form1.Hide 'Hide me
End Sub

Private Sub Command3_Click()
push = MsgBox("IF You Like This Code Please Vote For ME!", vbExclamation, "VOTE FOR JUSTIN!")
Unload Me ' shot me
End ' kill program
End Sub

Private Sub Command4_Click()
On Error Resume Next
Command4.Caption = GetIniData("Button2", "Button2", "Button2", App.Path & "/Button.ini") 'Insert Command Caption!
Shell GetIniData("Button2", "path", "", App.Path & "/Button.ini") 'Runs EXE file
End Sub

Private Sub Command5_Click()
On Error Resume Next
Command5.Caption = GetIniData("Button3", "Button3", "Button3", App.Path & "/Button.ini") 'Insert Command Caption!
Shell GetIniData("Button3", "path", "", App.Path & "/Button.ini") 'Runs EXE file
End Sub

Private Sub Command6_Click()
Unload Me
Frmmain.Show
End Sub

Private Sub Form_Load()
If Form1.MaxButton Then
push = MsgBox("IF You Like This Code Please Vote For ME!", vbExclamation, "VOTE FOR JUSTIN!")
End If
Command1.Caption = GetIniData("Button1", "Button1", "Button1", App.Path & "/Button.ini") 'Load Command Captions!
Command4.Caption = GetIniData("Button2", "Button2", "Button2", App.Path & "/Button.ini") 'Load Command Captions!
Command5.Caption = GetIniData("Button3", "Button3", "Button3", App.Path & "/Button.ini") 'Load Command Captions!
End Sub

Private Sub mnuabout_Click()
Form1.Hide
Form3.Show

End Sub

Private Sub mnuexit_Click()
Unload Form3 ' kill open forms
Unload Form2 ' kill open forms
Unload Me    ' kill me
End          ' kill program
End Sub

Private Sub mnuhow_Click()
Form1.Hide
Form4.Show
End Sub
