VERSION 5.00
Begin VB.Form zcjm 
   Caption         =   "你能想起你的账号吗？"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   10875
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox yhm 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox mm 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   5055
   End
   Begin VB.CommandButton zc 
      Caption         =   "注册"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton tc 
      Caption         =   "退出"
      Height          =   855
      Left            =   6480
      TabIndex        =   0
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   855
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   855
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "zcjm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tc_Click()
zcjm.Hide
End Sub

Private Sub zc_Click()
yh = yhm.Text
zz = yhm.Text
zm = mm.Text
Open App.Path & "\game\data\yhdata.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, yh
        If yh = zz Then
            GoTo find
        End If
Loop
Close #1
基础hp = 100
基础hpmax = 90
基础phymax = 90
hp = 100
hpmax = 100
att = 10
def = 0
health = 20
gold = 0
phy = 90
phymax = 90
game_day = 1
game_km = 0
eat = 100
eatmax = 100
gamemy = 1
gamemg = 1
gameyao_huifu = 30

Call loadzz
Open App.Path & "\game\data\yhdata.txt" For Append As #1
Print #1, zz
Close #1
GoTo ok
find:
MsgBox "你来晚了awa", , "账号"
Close #1
GoTo find2
ok:
MsgBox "已注册", , "注册"
find2:
End Sub
