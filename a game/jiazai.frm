VERSION 5.00
Begin VB.Form loading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "loading"
   ClientHeight    =   7155
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "jiazai.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7155
   ScaleWidth      =   13050
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox gamesjp 
      AutoSize        =   -1  'True
      Height          =   9780
      Left            =   0
      Picture         =   "jiazai.frx":0CCA
      ScaleHeight     =   9720
      ScaleWidth      =   17280
      TabIndex        =   0
      Top             =   0
      Width           =   17340
      Begin VB.Timer jz 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   600
         Top             =   3000
      End
      Begin VB.Timer jz2 
         Interval        =   200
         Left            =   600
         Top             =   2640
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   840
         Top             =   5040
         Width           =   11775
      End
      Begin VB.Label jm 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   8895
      End
      Begin VB.Label jdt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   5160
         Width           =   11535
      End
      Begin VB.Label bfb 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   4680
         Width           =   1935
      End
   End
End
Attribute VB_Name = "loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Randomize
sjp = Int(Rnd * 2 + 1)
'MsgBox sjp
If sjp = 2 Then
gamesjp.Picture = LoadPicture(App.Path & "\game\picture\bz\3.jpg")
End If
jz.Enabled = True
jz2.Enabled = True
Call loadget
End Sub


Private Sub jz_Timer()
Static a As Integer
Static b As Variant
b = "████████████████████████████████████████████████████████████████"
a = a + 1
jdt.Caption = jdt.Caption & Mid(b, a, 2)
If jdt.Caption = b Then
    game.Show
    jz.Enabled = False
    loading.Hide
End If
End Sub

Private Sub jz2_Timer()
Static a1 As Integer
Static b1 As Variant
b1 = "Loading..."
a1 = a1 + 1
jm.Caption = jm.Caption & Mid(b1, a1, 1)
If jm.Caption = b1 Then
    jz2.Enabled = False
End If
End Sub

