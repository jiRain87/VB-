VERSION 5.00
Begin VB.Form game 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "game"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "game.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7155
   ScaleWidth      =   11895
   StartUpPosition =   1  '����������
   Begin VB.TextBox gamebug 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9840
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.PictureBox xtp 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   7320
      MouseIcon       =   "game.frx":0CCA
      Picture         =   "game.frx":1594
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   22
      Top             =   5400
      Width           =   3015
   End
   Begin VB.PictureBox xzd 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1680
      MouseIcon       =   "game.frx":23FE
      Picture         =   "game.frx":2CC8
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   21
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Timer gamesjsj 
      Interval        =   10
      Left            =   11400
      Top             =   1200
   End
   Begin VB.PictureBox xsz 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5880
      Picture         =   "game.frx":3A97
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   18
      Top             =   6000
      Width           =   3015
   End
   Begin VB.PictureBox xsave 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3000
      Picture         =   "game.frx":48C6
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   16
      Top             =   6000
      Width           =   3015
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   3000
         ScaleHeight     =   375
         ScaleWidth      =   15
         TabIndex        =   17
         Top             =   120
         Width           =   15
      End
   End
   Begin VB.PictureBox xtx 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3000
      Picture         =   "game.frx":58B1
      ScaleHeight     =   615
      ScaleWidth      =   2895
      TabIndex        =   15
      Top             =   6600
      Width           =   2895
   End
   Begin VB.PictureBox xzz 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5880
      Picture         =   "game.frx":66CA
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   14
      Top             =   6600
      Width           =   3015
   End
   Begin VB.PictureBox xts 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      Picture         =   "game.frx":7494
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   13
      Top             =   6600
      Width           =   3015
   End
   Begin VB.PictureBox xbb 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8880
      Picture         =   "game.frx":82F5
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   12
      Top             =   6000
      Width           =   3015
   End
   Begin VB.PictureBox xsj 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   8880
      Picture         =   "game.frx":90E3
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   11
      Top             =   6600
      Width           =   3015
   End
   Begin VB.PictureBox xqj 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      MouseIcon       =   "game.frx":9F89
      Picture         =   "game.frx":A853
      ScaleHeight     =   615
      ScaleWidth      =   3015
      TabIndex        =   10
      Top             =   6000
      Width           =   3015
   End
   Begin VB.Label sjxxk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   20
      Top             =   3120
      Width           =   7455
   End
   Begin VB.Label zdxxk 
      BackStyle       =   0  'Transparent
      Height          =   3975
      Left            =   960
      TabIndex        =   19
      Top             =   1320
      Width           =   9735
   End
   Begin VB.Line Line12 
      Visible         =   0   'False
      X1              =   8919
      X2              =   8919
      Y1              =   6000
      Y2              =   7200
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   5946
      X2              =   5946
      Y1              =   6000
      Y2              =   7200
   End
   Begin VB.Line Line10 
      Visible         =   0   'False
      X1              =   2973
      X2              =   2973
      Y1              =   6000
      Y2              =   7200
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   0
      X2              =   11880
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   0
      X2              =   11880
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label gamephy 
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label gameeat 
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label xxx 
      Caption         =   "��ʱû��"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label gamegold 
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label gamedef 
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label gameatt 
      Height          =   255
      Left            =   7320
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line7 
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Label gamehealth 
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label gamehp 
      Height          =   255
      Left            =   9720
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label gamekm 
      Caption         =   "·��:???"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label gameday 
      Caption         =   "����:???"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line6 
      X1              =   7200
      X2              =   7200
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   9600
      X2              =   9600
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderStyle     =   3  'Dot
      X1              =   1800
      X2              =   11880
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   1800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   1800
      X2              =   1800
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   11880
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
game_bug = KeyCode

End Sub
Private Sub form_activate()
Call nagame
End Sub
Private Sub Form_Load()
'Call shuxing
gameday.Caption = "����:" & game_day
gamekm.Caption = "·��:" & game_km & "km"
gamegold.Caption = "���:" & gold
gamephy.Caption = "����:" & phy & "/" & phymax
gameatt.Caption = "����:" & att
gamedef.Caption = "����:" & def
gamehp.Caption = "����ֵ:" & hp & "/" & hpmax
gamehealth.Caption = "����:" & health
gameeat.Caption = "��ʳ��:" & eat & "/" & eatmax
gamesjsj.Enabled = False
xzd.Enabled = False: xzd.Visible = False
xtp.Enabled = False: xtp.Visible = False

gamebug.Enabled = False: gamebug.Visible = False
Call nagame
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Cancel = 0 Then
Cancel = 1
gameend = MsgBox("ȷ��Ҫ�˳���?", vbYesNo, "�˳���Ϸ")
End If
If gameend = 6 Then
End
ElseIf gameend = 7 Then
Cancel = 1
End If
End Sub

Private Sub gamebug_Change()
If gamebug.Text = "reload" Then
����hp = 100
����hpmax = 90
����phymax = 90
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
gamebug.Text = ""
Call nagame
ElseIf gamebug.Text = "hp" Then
hp = hpmax
End If
Call nagame
End Sub

Private Sub gamesjsj_Timer()
Static xxnum As Integer
Static zxx As Variant
Static wancheng As Integer
If wancheng = 1 Then
    sjxxk.Caption = ""
    xxnum = 0 '��Ϊ�ᱣ�� �������ó�ʼֵ
ElseIf wancheng = 0 Then
End If

If sjsj <= 5 Then
    If sjsjwp > 0 And sjsjwp <= 5 Then
        zxx = "��ȡ����X" & sjsjwps1 & "ľ��X" & sjsjwps2
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
        End If
    ElseIf sjsjwp > 5 And sjsjwp <= 9 Then
        zxx = "��ȡҩ��X" & sjsjwps1
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
        End If
    ElseIf sjsjwp > 9 And sjsjwp <= 11 Then
        zxx = "��ȡ޹�²�X" & sjsjwps1
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
        End If
    End If
ElseIf sjsj > 5 And sjsj <= 8 Then
        zxx = "��ʲô��û�ҵ�"
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
        End If
ElseIf sjsj > 9 And sjsj <= 10 Then
    If sjsjgw = 1 Then
        zxx = "������������"
        If xxnum >= Len(zxx) Then '==
        xxnum = 0 ''''''''''''''''''''==
        gamemake.makezs.Caption = "" '==
        End If '��ֹ��������Ϣ����==
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
        End If
    ElseIf sjsjgw = 0 Then
        zxx = "������������ �������ˡ�����"
        If xxnum >= Len(zxx) Then '==
        xxnum = 0 ''''''''''''''''''''==
        gamemake.makezs.Caption = "" '==
        End If '��ֹ��������Ϣ����==
        xxnum = xxnum + 1
        sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If sjxxk.Caption = zxx Then
            gamesjsj.Enabled = False
            wancheng = 1
            Call zdexit
        End If
    End If
End If
    
If sjsjwp > 100 And sjsjwp <= 200 Then
    zxx = "��Ĺ���������һ�� ����������һ��"
    xxnum = xxnum + 1
    sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
    wancheng = 0 '����Ϣû�������ʱ ��ɵ���0 ��ֹ����
    If sjxxk.Caption = zxx Then
        gamesjsj.Enabled = False
        wancheng = 1 '��Ϣ������
    End If
ElseIf sjsj > 300 Then
    zxx = "��Ĺ���������һ�� ����������һ��"
    xxnum = xxnum + 1
    wancheng = 0
    sjxxk.Caption = sjxxk.Caption & Mid(zxx, xxnum, 1)
    If sjxxk.Caption = zxx Then
        gamesjsj.Enabled = False
        wancheng = 1
    End If
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub xbb_Click()
zzbb = 1
gamemake.Show
game.Hide
End Sub

Private Sub xqj_Click()
game_km = game_km + 10
sjsj = Int((10 - 1 + 1) * Rnd + 1)
'MsgBox sjsj
If sjsj <= 5 Then '�ҵ���Ʒ
    sjsjwp = Int((11 - 1 + 1) * Rnd + 1) '�����Ʒ����
    'MsgBox sjsjwp
    If sjsjwp > 0 And sjsjwp <= 5 Then
        sjsjwps1 = Int((3 - 1 + 1) * Rnd + 1) '�����Ʒ����
        gameym = gameym + sjsjwps1
        sjsjwps2 = Int((3 - 1 + 1) * Rnd + 1)
        gamemc = gamemc + sjsjwps2
    ElseIf sjsjwp > 5 And sjsjwp <= 9 Then
        sjsjwps1 = Int((2 - 1 + 1) * Rnd + 1)
        gameyc = gameyc + sjsjwps1
    ElseIf sjsjwp > 9 And sjsjwp <= 11 Then
        sjsjwps1 = Int((3 - 1 + 1) * Rnd + 1)
        gamexyc = gamexyc + sjsjwps1
    End If
    gameday.Caption = "����:" & game_day '�����￪ʼ
    gamekm.Caption = "·��:" & game_km & "km"
    gamegold.Caption = "���:" & gold
    gamephy.Caption = "����:" & phy & "/" & phymax
    gameatt.Caption = "����:" & att
    gamedef.Caption = "����:" & def
    gamehp.Caption = "����ֵ:" & hp & "/" & hpmax
    gamehealth.Caption = "����:" & health
    gamesjsj.Enabled = True '��������� ����Ҫ ������
ElseIf sjsj > 5 And sjsj <= 8 Then '��ɶҲû�ҵ�
    'att = att - 1
    'def = def - 1
    gameday.Caption = "����:" & game_day
    gamekm.Caption = "·��:" & game_km & "km"
    gamegold.Caption = "���:" & gold
    gamephy.Caption = "����:" & phy & "/" & phymax
    gameatt.Caption = "����:" & att
    gamedef.Caption = "����:" & def
    gamehp.Caption = "����ֵ:" & hp & "/" & hpmax
    gamehealth.Caption = "����:" & health
    gamesjsj.Enabled = True
ElseIf sjsj > 9 And sjsj <= 10 Then 'ս��ģʽ
    sjsjgw = Int((1 - 0 + 1) * Rnd + 0)
    MsgBox sjsjgw
    Call zdopen
    gamesjsj.Enabled = True
End If
End Sub

Private Sub xzz_Click()
zzbb = 0
gamemake.Show
Call newpage
Call newpagegreen
'ˢ������
game.Hide
End Sub

