VERSION 5.00
Begin VB.Form login 
   Caption         =   "µ«¬º"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10860
   StartUpPosition =   1  'À˘”–’ﬂ÷––ƒ
   Begin VB.CommandButton zc 
      Caption         =   "◊¢≤·"
      Height          =   855
      Left            =   6240
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton dl 
      Caption         =   "µ«¬º"
      Height          =   855
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox mm 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   5055
   End
   Begin VB.TextBox yhm 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "√‹¬Î"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "”√ªß√˚"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dl_Click()
m = mm.Text
z = yhm.Text
If m = "" And z = "" Then
    MsgBox "«Î ‰»Î’À∫≈∫Õ√‹¬Î", , "µ«¬º"
    Exit Sub
End If
If z = "" And m <> "" Then
    MsgBox "«Î ‰»Î’À∫≈", , "µ«¬º"
    Exit Sub
ElseIf z <> "" And m = "" Then
    MsgBox "«Î ‰»Î√‹¬Î", , "µ«¬º"
    Exit Sub
End If
Open App.Path & "\game\data\yhdata.txt" For Input As #1
Do Until EOF(1)
    Line Input #1, yh
    If yh = z Then
        Close #1
        GoTo ok
    End If
Loop
Close #1
GoTo find
ok:
Open App.Path & "\game\saves\save" & z & ".txt" For Random As #1
Get #1, 1, zz
Get #1, 2, zm
Close #1
If zz = z And zm = m Then
start.Show
login.Hide
MsgBox "“—µ«¬º", , "µ«¬º"
yhm = zz
Else
find:
MsgBox "’À∫≈ªÚ√‹¬Î¥ÌŒÛ!", , "µ«¬º"
End If
End Sub

Private Sub zc_Click()
zcjm.Show
End Sub
