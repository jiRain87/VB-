VERSION 5.00
Begin VB.Form start 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "game"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11145
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox gamestartbz 
      Height          =   6135
      Left            =   0
      Picture         =   "start.frx":0000
      ScaleHeight     =   6075
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      Begin VB.CommandButton gamestart 
         Caption         =   "开始游戏"
         Height          =   855
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton tc 
         Caption         =   "退出"
         Height          =   855
         Left            =   3960
         TabIndex        =   3
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton config 
         Caption         =   "设置"
         Height          =   855
         Left            =   3960
         TabIndex        =   2
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label yhm 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub gamestart_Click()
loading.Show
start.Hide
End Sub

Private Sub tc_Click()
End
End Sub

Private Sub Form_Load()
yhm.Caption = "当前账号:" & zz
End Sub
