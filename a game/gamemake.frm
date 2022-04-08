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
   StartUpPosition =   1  '所有者中心
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
         Caption         =   "上一页"
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
         Caption         =   "下一页"
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
If zzbb = 0 Then '制造初始数值

    makezs.Caption = ""
    wppage = 1 '刷新初始数值
    gamemys = 5 + gamemy * 5
    gamemgs = 5 + gamemg * 5
    Call newpage
    Call newpagegreen
ElseIf zzbb = 1 Then
    makezs.Caption = ""
    wppage = 1 '刷新初始数值
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
        gameke1(0).Caption = "药"
        ke1.Caption = gameyc & "/2药草" & vbCrLf & "可以恢复30点血量，必要的补给品" & vbCrLf & "当前拥有" & gameyao
        gameke1(1).Caption = "熟肉"
        ke2.Caption = gamerou & "/1 肉" & gamemc & "/2 木材" & vbCrLf & "可以恢复70点饱食度 十分好吃的野味！" & vbCrLf & "当前拥有" & gamesrou
        gameke1(2).Caption = "麻衣LV" & gamemy
        ke3.Caption = gameym & "/" & gamemys & "亚麻" & vbCrLf & "每级提供10点防御和百分之五的格挡" _
        & vbCrLf & "当前提供" & gamemy * 5 & "的格挡几率"
        gameke1(3).Caption = "木棍LV" & gamemg
        ke4.Caption = gamemc & "/" & gamemgs & "木棍" & vbCrLf & "每级提供10点攻击和百分之五的暴击" _
        & vbCrLf & "当前提供" & gamemg * 5 & "的暴击几率"

        spage.Enabled = False: spage.Visible = False
        xpage.Enabled = True: xpage.Visible = True
    ElseIf wppage = 2 Then
        gameke1(0).Caption = "未增加"
        ke1.Caption = gameyc & "/2未增加" & Asc(13) & "未增加"
        gameke1(1).Caption = "未增加"
        ke2.Caption = gamerou & "/1 未增加" & gamemc & "/2 未增加" & Asc(13) & "未增加"
        gameke1(2).Caption = "未增加" & gamemy
        ke3.Caption = gameym & "/" & gamemys & Asc(13) & "未增加" _
        & Asc(13) & "未增加" & gamemy * 5 & "未增加"
        gameke1(3).Caption = "未增加" & gamemg
        ke4.Caption = gamemc & "/" & gamemgs & Asc(13) & "未增加五的暴击" _
        & Asc(13) & "当未增加供" & gamemg * 5 & "未增加率"
    
        spage.Enabled = True: spage.Visible = True
        xpage.Enabled = False: xpage.Visible = False
    End If
    Call newpagegreen '刷新信息
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

Private Sub maketime_Timer() '显示制造信息
Static xxnum As Integer
Static zxx As Variant
Static wancheng As Integer
If zzbb = 0 Then
    If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '因为会保存 所以重置初始值
    ElseIf wancheng = 0 Then
    End If
    If wppage = 1 Then
        If gameke1(0).Caption = "药" And gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
            zxx = "制造了一个" & gameke1(0).Caption
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '防止出现无信息问题
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' 刷新数据
            End If
        ElseIf gameke1(1).Caption = "熟肉" And gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
            zxx = "制造了一个" & gameke1(1).Caption
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '防止出现无信息问题
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' 刷新数据
            End If
        ElseIf gameke1(2).Caption = "麻衣LV" & gamemy - 1 And gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
            zxx = "制造了一个" & "麻衣LV" & gamemy
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '防止出现无信息问题
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' 刷新数据
            End If
        ElseIf gameke1(3).Caption = "木棍LV" & gamemg - 1 And gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
            zxx = "制造了一个" & "木棍LV" & gamemg
            If xxnum >= Len(zxx) Then
            xxnum = 0
            gamemake.makezs.Caption = ""
            End If '防止出现无信息问题
            xxnum = xxnum + 1
            makezs.Caption = makezs.Caption & Mid(zxx, xxnum, 1)
            wancheng = 0
            If makezs.Caption = zxx Then
                maketime.Enabled = False
                wancheng = 1
                Call newpage
                Call newpagegreen ' 刷新数据
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
