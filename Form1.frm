VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snake"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1830
   DrawWidth       =   5
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   122
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   720
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1440
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Score ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Pnt
  x As Integer
  y As Integer
End Type

Dim x As Integer, y As Integer
Dim snake(300) As Pnt, ns As Integer
Dim score As Integer

Private Sub cmdstart_Click()
Form1.Cls
cmdstart.Visible = False
Label1.Caption = "R"
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
ns = 4
For i = 0 To 4
  snake(i).x = 20 + 5 * i: snake(i).y = 30
Next i
' draw snake
For i = 0 To ns
  PSet (snake(i).x, snake(i).y), vbBlack
Next i
Line (5, 25)-(115, 115), vbBlue, B

x = 40: y = 30
score = 0
lblScore.Caption = 0
End Sub

Private Sub Form_Activate()
Line (5, 25)-(115, 115), vbBlue, B

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 37: Label1.Caption = "L"
  Case 38: Label1.Caption = "U"
  Case 39: Label1.Caption = "R"
  Case 40: Label1.Caption = "D"
  Case 27: End
End Select
End Sub

Private Sub Timer1_Timer()
' draw snake
PSet (snake(0).x, snake(0).y), &H8000000F  ' clear the tail
Call Check
For i = 0 To ns - 1
  snake(i).x = snake(i + 1).x
  snake(i).y = snake(i + 1).y
Next i
snake(ns).x = x: snake(ns).y = y
PSet (x, y), vbBlack
End Sub

Private Sub Timer2_Timer()
' increase the length of snake
If ns = 300 Then Exit Sub
Call Check
ns = ns + 1
snake(ns).x = x: snake(ns).y = y
PSet (x, y), vbBlack
End Sub

Private Sub Check()
Select Case Label1.Caption
  Case "U"
    y = y - 5
  Case "D"
    y = y + 5
  Case "L"
    x = x - 5
  Case "R"
    x = x + 5
End Select
' check ...
If Point(x, y) = vbBlack Then
  Timer1.Enabled = False
  Timer2.Enabled = False
  Timer3.Enabled = False
  If vbYes = MsgBox(" Oh ...Sorry" + vbCrLf + "Would you like to try again?", vbYesNo, "Game over ...") Then
    ' retry ...
    cmdstart.Visible = True
    Exit Sub
  Else
    End
  End If
End If
If Point(x, y) = vbRed Then
  score = score + 1
  lblScore = Val(score)
End If
If x = 115 Then x = 10
If x = 5 Then x = 110
If y = 25 Then y = 110
If y = 115 Then y = 30

End Sub

Private Sub Timer3_Timer()
PSet (Int(Rnd * 21) * 5 + 10, Int(Rnd * 17) * 5 + 30), vbRed
End Sub
