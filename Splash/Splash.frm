VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3960
   ClientLeft      =   4500
   ClientTop       =   3840
   ClientWidth     =   4980
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      Height          =   195
      Left            =   330
      TabIndex        =   0
      Top             =   3600
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   3960
      Left            =   0
      Picture         =   "Splash.frx":0000
      Top             =   0
      Width           =   4980
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Splash.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Splash = Nothing
End Sub

