VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraLogin 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load Old Game"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdPlayNew 
         Caption         =   "Play New Account"
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtAccount 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Account Name"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
Dim i As String

If Dir(App.Path & "\Accounts\Player1\") <> "" Then
    frmMain.Show
    frmLogin.Hide


Open App.Path & "\Accounts\Player1\Name.txt" For Input As #1
Input #1, i
Close #1
frmMain.lblName = i

Open App.Path & "\Accounts\Player1\Cards.txt" For Input As #1
Input #1, i
Close #1
frmMain.lblCards = i

Open App.Path & "\Accounts\Player1\Def.txt" For Input As #1
Input #1, i
Close #1
frmMain.lblDefense = i

Open App.Path & "\Accounts\Player1\Str.txt" For Input As #1
Input #1, i
Close #1
frmMain.lblStrength = i

Open App.Path & "\Accounts\Player1\Gold.txt" For Input As #1
Input #1, i
Close #1
frmMain.lblGold = i

End If

End Sub

Private Sub cmdPlayNew_Click()

If txtAccount.Text = "" Then
    MsgBox ("You need an account name derp.")
End If

If txtAccount.Text <> "" Then

frmMain.Show
frmLogin.Hide

With frmMain
.lblName.Caption = txtAccount.Text
.lblCards = "5"
.lblDefense = "1"
.lblStrength = "1"
.lblGold = "1"
End With

End If

End Sub

Private Sub Form_Load()

Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub
