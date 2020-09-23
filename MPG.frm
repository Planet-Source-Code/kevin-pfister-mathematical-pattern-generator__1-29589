VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mathematical Pattern Generator"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H0000FFFF&
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FFFF&
      ForeColor       =   &H0000FFFF&
      Height          =   11000
      Left            =   0
      ScaleHeight     =   10995
      ScaleWidth      =   16005
      TabIndex        =   0
      Top             =   0
      Width           =   16000
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   120
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+---------------------------------------------------+
'             Mathematical Pattern Generatator
'                 Created by Kevin Pfister
'+---------------------------------------------------+

Dim ctrX%, ctrY%
Const col% = 0
Dim a, b, c
Dim choss%
Dim x, y, xn
Dim n&, m&

Private Sub Form_Load()
Randomize
'Below picks the size and shape of the pattern
a = Int(Rnd * 100 + 1)
b = Int(Rnd * 100 + 1)
c = Int(Rnd * 100 + 1)
'Below is the patterns position on the screen
ctrX% = 200
ctrY% = 150
Timer1.Enabled = True
End Sub

Private Sub Picture1_Click()
End
End Sub

Private Sub Timer1_Timer()
        If x > 0 Then
            xn = y - Sqr(Abs(b * x - c))
        ElseIf x < 0 Then
            xn = y + Sqr(Abs(b * x - c))
        Else
            xn = y
        End If
        y = a - x
        x = xn
        Picture1.Line ((Int(x) + ctrX%) * 10, (Int(y) + ctrY%) * 10)-((Int(x) + ctrX% + 1) * 10, (1 + Int(y) + ctrY%) * 10)
        n& = n& + 1
End Sub
