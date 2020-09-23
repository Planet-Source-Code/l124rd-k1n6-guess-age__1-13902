VERSION 5.00
Begin VB.Form frmGuess 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Guess Age"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5625
   Icon            =   "Guessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   5
      Left            =   3240
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   4
      Left            =   2640
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   3
      Left            =   2040
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   2
      Left            =   1440
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   1
      Left            =   840
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Pic 
      Enabled         =   0   'False
      Height          =   6375
      Index           =   0
      Left            =   240
      ScaleHeight     =   6315
      ScaleWidth      =   315
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton btnAge 
      Caption         =   "&See Age"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton btnBegin 
      Caption         =   "&Begin"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   5
      Left            =   3360
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   4
      Left            =   2760
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   1560
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   960
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   0
      Left            =   360
      Shape           =   2  'Oval
      Top             =   6600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "AGE:"
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Make click on the columns where you see your age then click on See Age button. ^_^ \/"
      Height          =   1695
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu men 
      Caption         =   "&Menu"
      Begin VB.Menu sal 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu ace 
      Caption         =   "A&bout ..."
   End
End
Attribute VB_Name = "frmGuess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edad, flg(5) As Integer

Sub inicia()
'Purpose: Cleans form

   Label2.Caption = "AGE: "
   edad = 0
   Label1.Visible = True
   For X = 0 To 5
      Pic(X).Cls
      flg(X) = 0
      Shape1(X).Visible = False
      Pic(X).Enabled = True
   Next X
   btnAge.Enabled = False

End Sub

Private Sub ace_Click()

   frmAbout.Show 1
End Sub

Sub nums()
'Purpose: Initialize array

   For X = 1 To 63
      Pic(0).Print X
      X = X + 1
   Next X
   
   For X = 2 To 63
      For Y = 1 To 2
         Pic(1).Print X
         X = X + 1
      Next Y
      X = X + 1
   Next X
   
   For X = 4 To 63
      For Y = 1 To 4
         Pic(2).Print X
         X = X + 1
      Next Y
      X = X + 3
   Next X
   
   For X = 8 To 63
      For Y = 1 To 8
         Pic(3).Print X
         X = X + 1
      Next Y
      X = X + 7
   Next X
   
   For X = 16 To 63
      For Y = 1 To 16
         Pic(4).Print X
         X = X + 1
      Next Y
      X = X + 15
   Next X
   
   For X = 32 To 63
      Pic(5).Print X
   Next X

End Sub

Private Sub btnAge_Click()
'Purpose: Finally says your age

   If edad = 1 Then
      Label2.Caption = "AGE: " & edad & " YEAR"
      Exit Sub
   End If
   Label2.Caption = "AGE: " & edad & " YEARS"
End Sub

Private Sub btnBegin_Click()
'Purpose: Cleans form and Initialize the array of numbers

   Call inicia
   Call nums
   
End Sub

Private Sub Pic_Click(Index As Integer)
'Purpose: Here's the secret, when the person clicks the columns the program is adding
'         the numbers what are at the top of the columns, if you know binary system
   
   btnAge.Enabled = True
   Select Case Index
   Case 0:
      If flg(0) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 1
      Shape1(0).Visible = True
      flg(0) = 1
   Case 1:
      If flg(1) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 2
      Shape1(1).Visible = True
      flg(1) = 1
   Case 2:
      If flg(2) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 4
      Shape1(2).Visible = True
      flg(2) = 1
   Case 3:
      If flg(3) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 8
      Shape1(3).Visible = True
      flg(3) = 1
   Case 4:
      If flg(4) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 16
      Shape1(4).Visible = True
      flg(4) = 1
   Case 5:
      If flg(5) = 1 Then
         MsgBox ("ALREADY CLICKED"), 16, "I HAVE NO BUGS"
         Call inicia
         Call nums
         Exit Sub
      End If
      edad = edad + 32
      Shape1(5).Visible = True
      flg(5) = 1
   End Select
   
End Sub

Private Sub sal_Click()
'Purpose: Ends program
   End
End Sub
