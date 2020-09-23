VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   2115
   ClientTop       =   915
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   9960
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Caption         =   "Your App here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1485
      Left            =   1350
      TabIndex        =   0
      Top             =   3030
      Width           =   7275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Created by Pradeep Singh
'On 13-Dec-2001(11:45PM)
'E-mail - pradeepsingh10@hotmail.com

Private Sub Form_Load()
Form1.Visible = False
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
Splash.Show
Splash.SetFocus
Call DelayTime(2) 'Delay 2 Sec
Form1.Visible = True
Splash.SetFocus
Call DelayTime(2) 'Delay 2 Sec
Unload Splash
Form1.SetFocus
End Sub
