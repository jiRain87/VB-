VERSION 5.00
Begin VB.Form gamemake 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   11745
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Height          =   21975
      Left            =   0
      Picture         =   "gamemake.frx":0000
      ScaleHeight     =   21915
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Timer maketime 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   10560
         Top             =   600
      End
      Begin VB.Label makezs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   0
         Width           =   5175
      End
      Begin VB.Label spage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��һҳ"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Label ke4 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   1095
         Left            =   4200
         TabIndex        =   10
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label ke3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   1095
         Left            =   4200
         TabIndex        =   9
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label ke2 
         BackStyle       =   0  'Transparent
         Caption         =   "label2"
         ForeColor       =   &H8000000F&
         Height          =   1095
         Left            =   4200
         TabIndex        =   8
         Top             =   2280
         Width           =   3135
      End
      Begin VB.Label ke1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   1095
         Left            =   4200
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Line mak4a 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line mak4b 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label gameke1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   6
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Line mak3a 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line mak3b 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label gameke1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   2
         Left            =   4200
         TabIndex        =   5
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Line mak2a 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line mak2b 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label gameke1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H8000000F&
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   4
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label gameke1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.Line mak1b 
         BorderColor     =   &H8000000B&
         X1              =   3960
         X2              =   7560
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line mak1a 
         BorderColor     =   &H00FFFFFF&
         X1              =   3960
         X2              =   7560
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label xpage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "��һҳ"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   7200
         TabIndex        =   2
         Top             =   6480
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   2805
         X2              =   8805
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   2800
         X2              =   8800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label tc 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "gamemake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub form_activate()
If zzbb = 0 Then '�����ʼ��ֵ

    makezs.Caption = ""
    wppage = 1 'ˢ�³�ʼ��ֵ
    gamemys = 5 + gamemy * 5
    gamemgs = 5 + gamemg * 5
    Call newpage
    Call newpagegreen
ElseIf zzbb = 1 Then
    makezs.Caption = ""
    wppage = 1 'ˢ�³�ʼ��ֵ
    Call newpagebb
    Call newpagegreenbb
End If
maketime.Enabled = False
End Sub

Private Sub form_load()
wppage = 1
If zzbb = 0 Then
    gamemys = 5 + gamemy * 5
    gamemgs = 5 + gamemg * 5
    wppage = 1
    'gameyc = 4
    'gamemc = 10
    'gameym = 10
    If wppage = 1 Then
        gameke1(0).Caption = "ҩ"
        ke1.Caption = gameyc & "/2ҩ��" & vbCrLf & "���Իָ�30��Ѫ������Ҫ�Ĳ���Ʒ" & vbCrLf & "��ǰӵ��" & gameyao
        gameke1(1).Caption = "����"
        ke2.Caption = gamerou & "/1 ��" & gamemc & "/2 ľ��" & vbCrLf & "���Իָ�70�㱥ʳ�� ʮ�ֺóԵ�Ұζ��" & vbCrLf & "��ǰӵ��" & gamesrou
        gameke1(2).Caption = "����LV" & gamemy
        ke3.Caption = gameym & "/" & gamemys & "����" & vbCrLf & "ÿ���ṩ10������Ͱٷ�֮��ĸ�" _
        & vbCrLf & "��ǰ�ṩ" & gamemy * 5 & "�ĸ񵲼���"
        gameke1(3).Caption = "ľ��LV" & gamemg
        ke4.Caption = gamemc & "/" & gamemgs & "ľ��" & vbCrLf & "ÿ���ṩ10�㹥���Ͱٷ�֮��ı���" _
        & vbCrLf & "��ǰ�ṩ" & gamemg * 5 & "�ı�������"

        spage.Enabled = False: spage.Visible = False
        xpage.Enabled = True: xpage.Visible = True
    ElseIf wppage = 2 Then
        gameke1(0).Caption = "δ����"
        ke1.Caption = gameyc & "/2δ����" & Asc(13) & "δ����"
        gameke1(1).Caption = "δ����"
        ke2.Caption = gamerou & "/1 δ����" & gamemc & "/2 δ����" & Asc(13) & "δ����"
        gameke1(2).Caption = "δ����" & gamemy
        ke3.Caption = gameym & "/" & gamemys & Asc(13) & "δ����" _
        & Asc(13) & "δ����" & gamemy * 5 & "δ����"
        gameke1(3).Caption = "δ����" & gamemg
        ke4.Caption = gamemc & "/" & gamemgs & Asc(13) & "δ������ı���" _
        & Asc(13) & "��δ���ӹ�" & gamemg * 5 & "δ������"
    
        spage.Enabled = True: spage.Visible = True
        xpage.Enabled = False: xpage.Visible = False
    End If
    Call newpagegreen 'ˢ����Ϣ
ElseIf zzbb = 1 Then
gameyao = 100
wppage = 1
Call newpagebb
Call newpagegreenbb
End If
End Sub

Private Sub gameke1_Click(Index As Integer)
wpins = Index
If zzbb = 0 Then
    Call makewpin
ElseIf zzbb = 1 Then
    Call usewpin
End If
End Sub

Private Sub maketime_Timer() '��ʾ������Ϣ
Static xxnum As Integer
Static zxx As Variant
Static wancheng As Integer
If zzbb = 0 Then
    If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '��Ϊ�ᱣ�� �������ó�ʼֵ
    ElseIf wancheng = 0 Then
    End If
    If wppage = 1 Then
        If gameke1(0).Caption = "ҩ" And gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
            zxx = "������һ��" & gameke1(0).Caption
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '��ֹ��������Ϣ����
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' ˢ������
            End If
        ElseIf gameke1(1).Caption = "����" And gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
            zxx = "������һ��" & gameke1(1).Caption
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '��ֹ��������Ϣ����
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' ˢ������
            End If
        ElseIf gameke1(2).Caption = "����LV" & gamemy - 1 And gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
            zxx = "������һ��" & "����LV" & gamemy
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '��ֹ��������Ϣ����
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' ˢ������
            End If
        ElseIf gameke1(3).Caption = "ľ��LV" & gamemg - 1 And gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
            zxx = "������һ��" & "ľ��LV" & gamemg
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '��ֹ��������Ϣ����
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' ˢ������
            End If
        End If
    End If
ElseIf zzbb = 1 Then
    Call usewpinzs
End If
End Sub

Private Sub spage_Click()
If zzbb = 0 Then
    wppage = wppage - 1
    Call newpage
    Call newpagegreen
ElseIf zzbb = 1 Then
    wppage = wppage - 1
    Call newpagebb
    Call newpagegreenbb
End If
End Sub

Private Sub tc_Click()
gamemake.Hide
game.Show
End Sub

Private Sub xpage_Click()
If zzbb = 0 Then
    wppage = wppage + 1
    Call newpage
    Call newpagegreen
ElseIf zzbb = 1 Then
    wppage = wppage + 1
    Call newpagebb
    Call newpagegreenbb
End If
End Sub
