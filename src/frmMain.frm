VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kingdom of Hearts"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraShop 
      Caption         =   "Shop"
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Width           =   7215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "x"
      Height          =   255
      Left            =   7080
      TabIndex        =   18
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraConstants 
      Caption         =   "Constants"
      Height          =   3495
      Left            =   7440
      TabIndex        =   14
      Top             =   120
      Width           =   3735
      Begin VB.Frame fraPlayerConst 
         Caption         =   "Common Player Constants"
         Height          =   3135
         Left            =   1200
         TabIndex        =   16
         Top             =   240
         Width           =   2415
         Begin VB.CheckBox chkInFight 
            Caption         =   "In a fight?"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   240
      End
      Begin VB.Label lblConstantTimer 
         Caption         =   "0"
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraElse 
      Height          =   1095
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   3015
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fraProfile 
      Height          =   2415
      Left            =   3240
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      Begin VB.Frame fraStats 
         Caption         =   "Stats"
         Height          =   2175
         Left            =   2160
         TabIndex        =   9
         Top             =   120
         Width           =   1455
         Begin VB.Label Label2 
            Caption         =   "Def:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Str:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lblStrength 
            Caption         =   "Strength: "
            Height          =   255
            Left            =   600
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblDefense 
            Caption         =   "Defense: "
            Height          =   255
            Left            =   600
            TabIndex        =   10
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdStats 
         Caption         =   "Stats"
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Cards:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Gold:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblCards 
         Caption         =   "Cards: "
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblGold 
         Caption         =   "Gold: "
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdFight 
      Caption         =   "Find a Fight"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdShop 
      Caption         =   "Shop"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Profile"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdProfile 
      Caption         =   "My Profile"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function SpawnNpc()
Dim chance As Long

chance = Int((50 * Rnd) + 1)

End Function

Private Sub cmdExit_Click()

End

End Sub

Private Sub CmdFight_Click()

If chkInFight.Value = 1 Then
    lblNews.Caption = "You are already in a fight."
Else
    frmFindingFight.Show
    frmMain.Hide
    frmFindingFight.lblTimer.Caption = 5
    frmFindingFight.Label1.Caption = "                                                                                                Finding a fight."
End If

End Sub

Private Sub cmdProfile_Click()

If fraProfile.Visible = True Then
    fraProfile.Visible = False
    fraStats.Visible = False
Else
    fraProfile.Visible = True
End If

End Sub

Private Sub cmdSave_Click()

If chkInFight.Value = 1 Then
    lblNews.Caption = "You need to be out of a fight to save!"
    Exit Sub
End If

If Dir(App.Path & "\Accounts\Player1\") = "" Then
    MkDir (App.Path & "\Accounts\Player1\")
End If

Open App.Path & "\accounts\Player1\name.txt" For Output As #1
Print #1, lblName.Caption
Close #1

Open App.Path & "\accounts\Player1\gold.txt" For Output As #1
Print #1, lblGold.Caption
Close #1

Open App.Path & "\accounts\Player1\cards.txt" For Output As #1
Print #1, lblCards.Caption
Close #1

Open App.Path & "\accounts\Player1\str.txt" For Output As #1
Print #1, lblStrength.Caption
Close #1

Open App.Path & "\accounts\Player1\def.txt" For Output As #1
Print #1, lblDefense.Caption
Close #1

End Sub

Private Sub cmdStats_Click()

If fraStats.Visible = True Then
    fraStats.Visible = False
Else
    fraStats.Visible = True
End If

End Sub

Private Sub Form_Load()

fraProfile.Visible = False
fraStats.Visible = False

With frmMain
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With

End Sub

Private Sub Label1_Click()

lblNews.Caption = "Strength is used for increasing the chance of getting a higher number."

End Sub

Private Sub Label2_Click()

lblNews.Caption = "Defense is used for decreasing the chance of losing a card."

End Sub

Private Sub lblDefense_Click()

lblNews.Caption = "Defense is used for decreasing the chance of losing a card."

End Sub

Private Sub lblStrength_Click()

lblNews.Caption = "Strength is used for increasing the chance of getting a higher number."

End Sub

Private Sub Timer1_Timer()

lblConstantTimer.Caption = lblConstantTimer.Caption + 1

If lblConstantTimer.Caption = 2 Then
    lblConstantTimer.Caption = 0
End If

End Sub
