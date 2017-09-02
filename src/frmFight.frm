VERSION 5.00
Begin VB.Form frmFight 
   BorderStyle     =   0  'None
   Caption         =   "Fight!"
   ClientHeight    =   4560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   3840
      TabIndex        =   28
      Top             =   1440
      Width           =   2055
      Begin VB.CommandButton cmdLeave 
         Caption         =   "leave (5 gold)"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmdTurn 
         Caption         =   "Next Phase"
         Height          =   495
         Left            =   480
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblEAttack 
         Caption         =   "0"
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblAttack 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Yours                  Enemy's"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   5880
      TabIndex        =   15
      Top             =   1440
      Width           =   3735
      Begin VB.Frame Frame2 
         Caption         =   "Stats"
         Height          =   2175
         Left            =   2160
         TabIndex        =   16
         Top             =   120
         Width           =   1455
         Begin VB.Label Label10 
            Caption         =   "Def:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "Str:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label LblEStrength 
            Caption         =   "Strength: "
            Height          =   255
            Left            =   600
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblEDefense 
            Caption         =   "Defense: "
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Evil Dude"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Cards:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Gold:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblECards 
         Caption         =   "Cards: "
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblEGold 
         Caption         =   "Gold: "
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame frmTurn 
      Caption         =   "Turn"
      Height          =   615
      Left            =   4440
      TabIndex        =   13
      Top             =   840
      Width           =   855
      Begin VB.Label lblTurn 
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraProfile 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3735
      Begin VB.Frame fraStats 
         Caption         =   "Stats"
         Height          =   2175
         Left            =   2160
         TabIndex        =   1
         Top             =   120
         Width           =   1455
         Begin VB.Label lblDefense 
            Caption         =   "Defense: "
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblStrength 
            Caption         =   "Strength: "
            Height          =   255
            Left            =   600
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Str:"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "Def:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   375
         End
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblGold 
         Caption         =   "Gold: "
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCards 
         Caption         =   "Cards: "
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Gold:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Cards:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblEnemyLevel 
      Alignment       =   2  'Center
      Caption         =   "ENEMY"
      Height          =   495
      Left            =   6360
      TabIndex        =   26
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "YOU"
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmFight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLeave_Click()
Dim possible As Boolean

If lblTurn.Caption = 1 Then
    If lblGold.Caption >= 5 Then
        MsgBox ("You leave the game.")
        frmMain.lblGold.Caption = frmMain.lblGold.Caption - 5
        possible = True
    End If
End If

If lblTurn.Caption = 2 Then
    If lblGold.Caption >= 3 Then
        MsgBox ("You leave the game.")
        frmMain.lblGold.Caption = frmMain.lblGold.Caption - 3
        possible = True
    End If
End If

If lblTurn.Caption = 3 Then
    If lblGold.Caption >= 1 Then
        MsgBox ("You leave the game.")
        frmMain.lblGold.Caption = frmMain.lblGold.Caption - 1
        possible = True
    End If
End If

If lblTurn.Caption > 3 Then
    MsgBox ("You leave the game.")
    possible = True
End If

If possible = True Then
    frmMain.lblCards.Caption = lblCards.Caption
    frmMain.Show
    frmFight.Hide
    frmMain.chkInFight.Value = 0
End If
        
End Sub

Private Sub cmdTurn_Click()
Dim damage As Long

lblTurn.Caption = lblTurn.Caption + 1

If lblTurn.Caption = 2 Then
    cmdLeave.Caption = "leave (3 gold)"
End If

If lblTurn.Caption = 3 Then
    cmdLeave.Caption = "leave (1 gold)"
End If

If lblTurn.Caption > 3 Then
    cmdLeave.Caption = "leave"
End If

damage = Int((100 * Rnd) + 1)
lblAttack = damage + lblStrength

damage = Int((100 * Rnd) + 1)
lblEAttack = damage + LblEStrength

If lblEAttack.Caption > lblAttack Then
    lblECards.Caption = lblECards.Caption + 1
    lblCards.Caption = lblCards.Caption - 1
End If

If lblAttack.Caption > lblEAttack Then
    lblECards.Caption = lblECards.Caption - 1
    lblCards.Caption = lblCards.Caption + 1
End If

    If lblECards.Caption = 0 Then
        MsgBox ("You win!")
        damage = Int((lblEGold.Caption * Rnd) + 1)
        frmMain.lblGold.Caption = frmMain.lblGold.Caption + damage
        frmMain.lblCards.Caption = lblCards.Caption
        
        frmMain.Show
        frmFight.Hide
        frmMain.chkInFight.Value = 0
    End If
    
    If lblCards.Caption = 0 Then
        damage = Int((lblGold.Caption * Rnd) + 1)
        MsgBox ("You lost all your cards! You can buy some at the store.")
        frmMain.lblGold.Caption = frmMain.lblGold.Caption - damage
        frmMain.lblCards.Caption = 0
        
        If frmMain.lblGold.Caption = 0 Then
            MsgBox ("We just realized you lost all your cards. Good job. You lose the entire game.")
            End
        End If
        
        frmMain.Show
        frmFight.Hide
        frmMain.chkInFight.Value = 0
    End If

End Sub

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

With frmFight
    .lblName = frmMain.lblName.Caption
    .lblGold = frmMain.lblGold.Caption
    .lblCards = frmMain.lblCards.Caption
    .lblStrength = frmMain.lblStrength.Caption
    .lblDefense = frmMain.lblDefense.Caption
End With

Call BufferEnemyLevel


' loading the enemy's stats
'With frmFight

    

End Sub

Public Function BufferEnemyLevel()
Dim level As Long

level = Int((100 * Rnd) + 1)

lblEGold.Caption = level

level = Int((100 * Rnd) + 1)

If level < 25 Then
    Call BufferWeakEnemy
    lblEnemyLevel.Caption = "Weak Enemy"
    Exit Function
End If

If level > 25 And level < 50 Then
    Call BufferMediumEnemy
    lblEnemyLevel.Caption = "Medium Enemy"
    Exit Function
End If

If level > 50 And level < 75 Then
    Call BufferHardEnemy
    lblEnemyLevel.Caption = "Hard Enemy"
    Exit Function
End If

If level > 75 Then
    Call BufferExtremeEnemy
    lblEnemyLevel.Caption = "Extreme Enemy"
    Exit Function
End If

End Function

Private Function BufferWeakEnemy()
Dim level As Long

level = Int((5 * Rnd) + 1)

lblECards.Caption = level
LblEStrength.Caption = 1
lblEDefense.Caption = 1

End Function

Private Function BufferMediumEnemy()
Dim level As Long

level = Int((10 * Rnd) + 1)
lblECards.Caption = level

level = Int((5 * Rnd) + 1)
LblEStrength.Caption = level

level = Int((5 * Rnd) + 1)
lblEDefense.Caption = level

End Function

Private Function BufferHardEnemy()
Dim level As Long

level = Int((20 * Rnd) + 1)
lblECards.Caption = level

level = Int((10 * Rnd) + 1)
LblEStrength.Caption = level

level = Int((10 * Rnd) + 1)
lblEDefense.Caption = level

End Function

Private Function BufferExtremeEnemy()
Dim level As Long

level = Int((25 * Rnd) + 1)
lblECards.Caption = level

level = Int((15 * Rnd) + 1)
LblEStrength.Caption = level

level = Int((15 * Rnd) + 1)
lblEDefense.Caption = level

End Function
