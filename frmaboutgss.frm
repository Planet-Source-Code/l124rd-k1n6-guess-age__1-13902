VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Guess Age"
   ClientHeight    =   2835
   ClientLeft      =   6450
   ClientTop       =   2220
   ClientWidth     =   5145
   ClipControls    =   0   'False
   Icon            =   "frmaboutgss.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1956.767
   ScaleMode       =   0  'User
   ScaleWidth      =   4831.421
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmaboutgss.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton btnOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   $"frmaboutgss.frx":0614
      Height          =   855
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmaboutgss.frx":06E5
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Adivina Edad"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   945
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.0.0"
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   930
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Integer
Dim yy As Integer

Private Sub btnOK_Click()
'Purpose: Unload frmAbout

  Unload Me
End Sub

Private Sub btnOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Purpose: Moves the button

   Randomize
   btnOK.Caption = "JEJEJE"
   xx = (Rnd * frmAbout.Width) - (btnOK.Width / 2) '5235
   yy = (Rnd * frmAbout.Height) - (btnOK.Height / 2) '3180
   btnOK.Left = xx
   btnOK.Top = yy
   
End Sub
