VERSION 5.00
Begin VB.Form frmFindingFight 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerFind 
      Interval        =   1000
      Left            =   5880
      Top             =   240
   End
   Begin VB.Label lblTimer 
      Caption         =   "5"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "                                                                                                Finding a fight."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmFindingFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

With frmFindingFight
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

timerFind.Enabled = True

End Sub

Private Sub timerFind_Timer()

lblTimer.Caption = lblTimer.Caption - 1

If lblTimer.Caption = 4 Then Label1.Caption = Label1.Caption + "."
If lblTimer.Caption = 3 Then Label1.Caption = Label1.Caption + "."
If lblTimer.Caption = 2 Then Label1.Caption = Label1.Caption + "."
If lblTimer.Caption = 1 Then Label1.Caption = "                                                                                                Fight Found."


If lblTimer.Caption = 0 Then
    frmFight.Show
    frmFindingFight.Hide
    frmMain.chkInFight.Value = 1
    timerFind.Enabled = False
    Call frmFight.BufferEnemyLevel
End If

End Sub
