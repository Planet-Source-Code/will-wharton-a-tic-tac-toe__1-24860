VERSION 5.00
Begin VB.Form ttt 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4575
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5940
   Icon            =   "ttt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   120
      Top             =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Player O"
      Height          =   1572
      Left            =   4080
      TabIndex        =   1
      Top             =   2640
      Width           =   1692
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1212
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Player X"
      Height          =   1572
      Left            =   4080
      TabIndex        =   0
      Top             =   840
      Width           =   1692
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1212
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   360
   End
   Begin VB.Image Image3 
      Height          =   1095
      Left            =   480
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Image Image9 
      Height          =   1095
      Left            =   2760
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Image Image8 
      Height          =   1215
      Left            =   2760
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image7 
      Height          =   1095
      Left            =   2760
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image6 
      Height          =   1095
      Left            =   1560
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   1215
      Left            =   1560
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   1560
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   1215
      Left            =   480
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   480
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image O9 
      Height          =   960
      Left            =   2880
      Picture         =   "ttt.frx":030A
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   960
   End
   Begin VB.Image O8 
      Height          =   960
      Left            =   2880
      Picture         =   "ttt.frx":3C4C
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   960
   End
   Begin VB.Image O7 
      Height          =   960
      Left            =   2880
      Picture         =   "ttt.frx":758E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   960
   End
   Begin VB.Image O6 
      Height          =   960
      Left            =   1680
      Picture         =   "ttt.frx":AED0
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   945
   End
   Begin VB.Image O5 
      Height          =   945
      Left            =   1680
      Picture         =   "ttt.frx":E812
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   945
   End
   Begin VB.Image O4 
      Height          =   960
      Left            =   1680
      Picture         =   "ttt.frx":12154
      Stretch         =   -1  'True
      Top             =   840
      Width           =   975
   End
   Begin VB.Image O2 
      Height          =   960
      Left            =   480
      Picture         =   "ttt.frx":15A96
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   960
   End
   Begin VB.Image O3 
      Height          =   960
      Left            =   480
      Picture         =   "ttt.frx":193D8
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   960
   End
   Begin VB.Image O1 
      Height          =   960
      Left            =   480
      Picture         =   "ttt.frx":1CD1A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   960
   End
   Begin VB.Image X6 
      Height          =   852
      Left            =   1680
      Picture         =   "ttt.frx":2065C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   972
   End
   Begin VB.Image X9 
      Height          =   855
      Left            =   2880
      Picture         =   "ttt.frx":23F9E
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   855
   End
   Begin VB.Image X3 
      Height          =   852
      Left            =   600
      Picture         =   "ttt.frx":278E0
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   852
   End
   Begin VB.Image X5 
      Height          =   975
      Left            =   1680
      Picture         =   "ttt.frx":2B222
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image X8 
      Height          =   972
      Left            =   2880
      Picture         =   "ttt.frx":2EB64
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   852
   End
   Begin VB.Image X2 
      Height          =   972
      Left            =   600
      Picture         =   "ttt.frx":324A6
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   852
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Image X7 
      Height          =   852
      Left            =   2880
      Picture         =   "ttt.frx":35DE8
      Stretch         =   -1  'True
      Top             =   960
      Width           =   852
   End
   Begin VB.Image X1 
      Height          =   852
      Left            =   600
      Picture         =   "ttt.frx":3972A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   852
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   3840
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image X4 
      Height          =   852
      Left            =   1680
      Picture         =   "ttt.frx":3D06C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   972
   End
   Begin VB.Line Line2 
      X1              =   2760
      X2              =   2760
      Y1              =   840
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1560
      Y1              =   840
      Y2              =   4200
   End
End
Attribute VB_Name = "ttt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sp1 As Boolean
Dim sp2 As Boolean
Dim p1p As Integer
Dim p2p As Integer

Function lösch()

O1.Visible = False
O2.Visible = False
O3.Visible = False
O4.Visible = False
O5.Visible = False
O6.Visible = False
O7.Visible = False
O8.Visible = False
O9.Visible = False

X1.Visible = False
X2.Visible = False
X3.Visible = False
X4.Visible = False
X5.Visible = False
X6.Visible = False
X7.Visible = False
X8.Visible = False
X9.Visible = False
End Function

Private Sub Command1_Click()
    MsgBox "Cats game", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    cat = cat + 1
    Timer2.Enabled = True
End Sub

Private Sub Image1_Click()

If sp1 = True And O1.Visible = False Then
    X1.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X1.Visible = False Then
    O1.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Form_Load()

sp1 = True
sp2 = False


X1.Visible = False
X2.Visible = False
X3.Visible = False
X4.Visible = False
X5.Visible = False
X6.Visible = False
X7.Visible = False
X8.Visible = False
X9.Visible = False

O1.Visible = False
O2.Visible = False
O3.Visible = False
O4.Visible = False
O5.Visible = False
O6.Visible = False
O7.Visible = False
O8.Visible = False
O9.Visible = False
End Sub

Private Sub Image2_Click()


If sp1 = True And O2.Visible = False Then
    X2.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X2.Visible = False Then
    O2.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image3_Click()

If sp1 = True And O3.Visible = False Then
    X3.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X3.Visible = False Then
    O3.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image4_Click()

If sp1 = True And O4.Visible = False Then
    X4.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X4.Visible = False Then
    O4.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image5_Click()

If sp1 = True And O5.Visible = False Then
    X5.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X5.Visible = False Then
    O5.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image6_Click()
If sp1 = True And O6.Visible = False Then
    X6.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X6.Visible = False Then
    O6.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image7_Click()
If sp1 = True And O7.Visible = False Then
    X7.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X7.Visible = False Then
    O7.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image8_Click()
If sp1 = True And O8.Visible = False Then
    X8.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X8.Visible = False Then
    O8.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub

Private Sub Image9_Click()
If sp1 = True And O9.Visible = False Then
    X9.Visible = True
    sp2 = True
    sp1 = False
    Exit Sub
End If

If sp2 = True And X9.Visible = False Then
    O9.Visible = True
    sp1 = True
    sp2 = False
    Exit Sub
End If

End Sub
Private Sub Timer1_Timer()
If sp1 = True Then ttt.MousePointer = 2
If sp2 = True Then ttt.MousePointer = 12
End Sub

Private Sub Timer2_Timer()

If X1.Visible = True And X2.Visible = True And X3.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X1.Visible = True And X4.Visible = True And X7.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X1.Visible = True And X5.Visible = True And X9.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X2.Visible = True And X5.Visible = True And X8.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X3.Visible = True And X6.Visible = True And X9.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X3.Visible = True And X5.Visible = True And X7.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X4.Visible = True And X5.Visible = True And X6.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If

If X7.Visible = True And X8.Visible = True And X9.Visible = True Then
    MsgBox "Player X wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p1p = p1p + 1
    Timer2.Enabled = True
End If


If O1.Visible = True And O2.Visible = True And O3.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O1.Visible = True And O4.Visible = True And O7.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O1.Visible = True And O5.Visible = True And O9.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O2.Visible = True And O5.Visible = True And O8.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O3.Visible = True And O6.Visible = True And O9.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O3.Visible = True And O5.Visible = True And O7.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O4.Visible = True And O5.Visible = True And O6.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

If O7.Visible = True And O8.Visible = True And O9.Visible = True Then
    MsgBox "Player O wins", vbOKOnly, "The End"
    Timer2.Enabled = False
    lösch
    p2p = p2p + 1
    Timer2.Enabled = True
End If

End Sub

Private Sub Timer3_Timer()

Label3.Caption = p1p
Label4.Caption = p2p
End Sub
