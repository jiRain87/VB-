Attribute VB_Name = "Module1"
Public m, z, zm, zz, yhm '用户与密码
Public hp%, hpmax%, eat%, eatmax%, att%, def%, health%, gold%, phy%, phymax%, game_day As Integer, game_km As Integer '定义游戏用户属性
'分别是 生命值 生命上限 饱食度 饱食度上限 攻击 防御 健康 金币 体力 体力上限 游戏天数 游戏路程
Public 基础hp, 基础hpmax, 基础att, 基础def, 基础phymax '定义基础属性值
Public gwhp, gwatt, gwdef '定义怪物属性
'分别是怪物 血量 攻击 防御
Public sjsj As Integer '事件
Public sjsjwp As Integer, sjsjwps1, sjsjwps2, sjsjwps3 As Integer '随机物品
'依次是物品事件 物品数量
Public sjsjgw As Integer '怪物类型
Public sjp As Integer '随机壁纸
Public gameend As Integer '退出时返回的数值
Public gameym%, gameyc%, gamexyc%, gamemc%, gamerou As Integer '定义材料属性全为整型
'从左到右分别为 亚麻 药草 薰衣草 木材 生肉
Public gameapple%, gamesrou%, gameyao%, gameyan As Integer '定义物品
'从左到右 苹果 熟肉 药 烟
Public gameyao_huifu As Integer '伤药恢复的hp
Public gamemg%, gamemy As Integer '定义装备
'为木棍 麻衣
Public gamemgs%, gamemys As Integer '定义升级所需数量
Public wppage%, wpins As Integer '制造页页数 物品数位置
Public zzbb As Integer '判断是制造还是背包 制造0 背包1
Public Function loadzz() '保存存档
Open App.Path & "\game\saves\save" & zz & ".txt" For Random As #1
Put #1, 1, zz
Put #1, 2, zm
Put #1, 3, hp
Put #1, 4, hpmax
Put #1, 5, att
Put #1, 6, def
Put #1, 7, health
Put #1, 8, gold
Put #1, 9, phy
Put #1, 10, phymax
Put #1, 11, game_day
Put #1, 12, game_km
Put #1, 13, eat
Put #1, 14, eatmax
Put #1, 15, 基础hp
Put #1, 16, 基础hpmax
Put #1, 17, 基础att
Put #1, 18, 基础def
Put #1, 19, 基础phymax
Put #1, 20, gameym
Put #1, 21, gameyc
Put #1, 22, gamexyc
Put #1, 23, gamemc
Put #1, 24, gamerou
Put #1, 25, gameapple
Put #1, 26, gamesrou
Put #1, 27, gameyao
Put #1, 28, gameyan
Put #1, 29, gameyao_huifu
Put #1, 30, gamemg
Put #1, 31, gamemy
Put #1, 32, gamemgs
Put #1, 33, gamemys
Close #1
End Function



Public Function loadget() '读取存档
Open App.Path & "\game\saves\save" & zz & ".txt" For Random As #1
Get #1, 1, zz
Get #1, 2, zm
Get #1, 3, hp
Get #1, 4, hpmax
Get #1, 5, att
Get #1, 6, def
Get #1, 7, health
Get #1, 8, gold
Get #1, 9, phy
Get #1, 10, phymax
Get #1, 11, game_day
Get #1, 12, game_km
Get #1, 13, eat
Get #1, 14, eatmax
Get #1, 15, 基础hp
Get #1, 16, 基础hpmax
Get #1, 17, 基础att
Get #1, 18, 基础def
Get #1, 19, 基础phymax
Get #1, 20, gameym
Get #1, 21, gameyc
Get #1, 22, gamexyc
Get #1, 23, gamemc
Get #1, 24, gamerou
Get #1, 25, gameapple
Get #1, 26, gamesrou
Get #1, 27, gameyao
Get #1, 28, gameyan
Get #1, 29, gameyao_huifu
Get #1, 30, gamemg
Get #1, 31, gamemy
Get #1, 32, gamemgs
Get #1, 33, gamemys
Close #1
End Function



Public Function shuxing() '未使用
att = gamemg * 10
lovemax = 基础lovemax + gamemy * 10
End Function



Public Sub nagame() '刷新界面\
If hp > hpmax Then
    hp = hpmax
End If
If phy > phymax Then
    phy = phymax
End If
game.gameday.Caption = "天数:" & game_day
game.gamekm.Caption = "路程:" & game_km & "km"
game.gamegold.Caption = "金币:" & gold
game.gamephy.Caption = "体力:" & phy & "/" & phymax
game.gameatt.Caption = "攻击:" & att
game.gamedef.Caption = "防御:" & def
game.gamehp.Caption = "生命值:" & hp & "/" & hpmax
game.gamehealth.Caption = "健康:" & health
End Sub



Public Function newpage() '刷新属性 以及翻页导致的变动
If wppage = 1 Then
    gamemake.gameke1(0).Caption = "药"
    gamemake.ke1.Caption = gameyc & "/2药草" & vbCrLf & "可以恢复" & gameyao_huifu & "点血量，必要的补给品" & vbCrLf & "当前拥有" & gameyao
    gamemake.gameke1(1).Caption = "熟肉"
    gamemake.ke2.Caption = gamerou & "/1 肉" & gamemc & "/2 木材" & vbCrLf & "可以恢复70点饱食度 十分好吃的野味！" & vbCrLf & "当前拥有" & gamesrou
    gamemake.gameke1(2).Caption = "麻衣LV" & gamemy
    gamemake.ke3.Caption = gameym & "/" & gamemys & "亚麻" & vbCrLf & "每级提供" & gamemy * 10 & "点防御和百分之五的格挡" _
    & vbCrLf & "当前提供" & gamemy * 5 & "的格挡几率"
    gamemake.gameke1(3).Caption = "木棍LV" & gamemg
    gamemake.ke4.Caption = gamemc & "/" & gamemgs & "木棍" & vbCrLf & "每级提供" & gamemg * 10 & "点攻击和百分之五的暴击" _
    & vbCrLf & "当前提供" & gamemg * 5 & "的暴击几率"

    gamemake.spage.Enabled = False: gamemake.spage.Visible = False
    gamemake.xpage.Enabled = True: gamemake.xpage.Visible = True
ElseIf wppage = 2 Then
    gamemake.gameke1(0).Caption = "未增加"
    gamemake.ke1.Caption = gameyc & "/2未增加" & Asc(13) & "未增加"
    gamemake.gameke1(1).Caption = "未增加"
    gamemake.ke2.Caption = gamerou & "/1 未增加" & gamemc & "/2 未增加" & Asc(13) & "未增加"
    gamemake.gameke1(2).Caption = "未增加" & gamemy
    gamemake.ke3.Caption = gameym & "/" & gamemys & Asc(13) & "未增加" _
    & Asc(13) & "未增加" & gamemy * 5 & "未增加"
    gamemake.gameke1(3).Caption = "未增加" & gamemg
    gamemake.ke4.Caption = gamemc & "/" & gamemgs & Asc(13) & "未增加五的暴击" _
    & Asc(13) & "当未增加供" & gamemg * 5 & "未增加率"
    
    gamemake.spage.Enabled = True: gamemake.spage.Visible = True
    gamemake.xpage.Enabled = False: gamemake.xpage.Visible = False
End If
End Function



Public Function newpagegreen() '刷新制造
If gameyc >= 2 Then
    gamemake.gameke1(0).ForeColor = &HFF00&
    gamemake.mak1a.BorderColor = &HFF00&: gamemake.mak1b.BorderColor = &HFF00& '如果符合条件 把颜色弄成绿色
ElseIf gameyc < 2 Then
    gamemake.gameke1(0).ForeColor = &H8000000F
    gamemake.mak1a.BorderColor = &H8000000F: gamemake.mak1b.BorderColor = &H8000000F '如果他又不符合了 把颜色弄成白色
End If
If gamerou >= 1 And gamemc >= 2 Then
    gamemake.gameke1(1).ForeColor = &HFF00&
    gamemake.mak2a.BorderColor = &HFF00&: gamemake.mak2b.BorderColor = &HFF00&
ElseIf gamerou < 1 And gamemc < 2 Then
    gamemake.gameke1(1).ForeColor = &H8000000F
    gamemake.mak2a.BorderColor = &H8000000F: gamemake.mak2b.BorderColor = &H8000000F
    
End If
If gameym >= gamemys Then
    gamemake.gameke1(2).ForeColor = &HFF00&
    gamemake.mak3a.BorderColor = &HFF00&: gamemake.mak3b.BorderColor = &HFF00&
ElseIf gameym < gamemys Then
    gamemake.gameke1(2).ForeColor = &H8000000F
    gamemake.mak3a.BorderColor = &H8000000F: gamemake.mak3b.BorderColor = &H8000000F
End If
If gamemc >= gamemgs Then
    gamemake.gameke1(3).ForeColor = &HFF00&
    gamemake.mak4a.BorderColor = &HFF00&: gamemake.mak4b.BorderColor = &HFF00&
ElseIf gamemc < gamemgs Then
    gamemake.gameke1(3).ForeColor = &H8000000F
    gamemake.mak4a.BorderColor = &H8000000F: gamemake.mak4b.BorderColor = &H8000000F
End If
End Function



Public Function makewpin() '制造物品
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "药" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        gameyc = gameyc - 2
        gameyao = gameyao + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(1).Caption = "熟肉" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        gamemc = gamemc - 2
        gamerou = gamerou - 1
        gamesrou = gamesrou + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(2).Caption = "麻衣LV" & gamemy And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
        gameym = gameym - gamemys '去掉代价
        gamemy = gamemy + 1 '增加等级
        hpmax = hpmax + 50
        hp = hp + 50 '拿走你应得的
        gamemys = 5 + gamemy * 5 '增加代价
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(3).Caption = "木棍LV" & gamemg And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
        gamemc = gamemc - gamemgs
        gamemg = gamemg + 1
        att = att + 10
        gamemgs = 5 + gamemg * 5
        gamemake.maketime.Enabled = True
    End If
    
End If
End Function


Public Function makewpinzs() '物品制造信息
If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '因为会保存 所以重置初始值
ElseIf wancheng = 0 Then
End If
If wppage = 1 Then
    If gameke1(0).Caption = "药" And gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        zxx = "制造了一个" & gameke1(0).Caption
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
End Function

Public Function newpagebb() ' 背包的翻页信息刷新
If wppage = 1 Then
    gamemake.gameke1(0).Caption = "药"
    gamemake.ke1.Caption = "恢复" & gameyao_huifu & "点生命" & vbCrLf & "当前拥有" & gameyao
    gamemake.gameke1(1).Caption = "熟肉"
    gamemake.ke2.Caption = "可以恢复70点饱食度 十分好吃的野味！" & vbCrLf & "当前拥有" & gamesrou
    gamemake.gameke1(2).Caption = "未有"
    gamemake.ke3.Caption = "暂未有"
    gamemake.gameke1(3).Caption = 暂未有
    gamemake.ke4.Caption = "暂未有"
    
    gamemake.spage.Enabled = False: gamemake.spage.Visible = False
    gamemake.xpage.Enabled = True: gamemake.xpage.Visible = True
ElseIf wppage = 2 Then
    gamemake.gameke1(0).Caption = "未增加"
    gamemake.ke1.Caption = "未增加" & Asc(13) & "未增加"
    gamemake.gameke1(1).Caption = "未增加"
    gamemake.ke2.Caption = "未增加/2 未增加" & Asc(13) & "未增加"
    gamemake.gameke1(2).Caption = "未增加"
    gamemake.ke3.Caption = "/" & Asc(13) & "未增加" _
    & Asc(13) & "未增加" & "未增加"
    gamemake.gameke1(3).Caption = "未增加"
    gamemake.ke4.Caption = "/" & Asc(13) & "未增加的暴击" _
    & Asc(13) & "当未增加供未增加率"
    
    gamemake.spage.Enabled = True: gamemake.spage.Visible = True
    gamemake.xpage.Enabled = False: gamemake.xpage.Visible = False
End If
End Function




Public Function newpagegreenbb() '背包有效性刷新
If gameyao > 0 Then
    gamemake.gameke1(0).ForeColor = &HFF00&
    gamemake.mak1a.BorderColor = &HFF00&: gamemake.mak1b.BorderColor = &HFF00& '如果符合条件 把颜色弄成绿色
ElseIf gameyao <= 0 Then
    gamemake.gameke1(0).ForeColor = &H8000000F
    gamemake.mak1a.BorderColor = &H8000000F: gamemake.mak1b.BorderColor = &H8000000F '如果他又不符合了 把颜色弄成白色
End If
If gamesrou > 0 Then
    gamemake.gameke1(1).ForeColor = &HFF00&
    gamemake.mak2a.BorderColor = &HFF00&: gamemake.mak2b.BorderColor = &HFF00&
ElseIf gamesrou <= 0 Then
    gamemake.gameke1(1).ForeColor = &H8000000F
    gamemake.mak2a.BorderColor = &H8000000F: gamemake.mak2b.BorderColor = &H8000000F
    
End If
If gameym >= 10000000 Then
    gamemake.gameke1(2).ForeColor = &HFF00&
    gamemake.mak3a.BorderColor = &HFF00&: gamemake.mak3b.BorderColor = &HFF00&
ElseIf gameym < -1000000000 Then
    gamemake.gameke1(2).ForeColor = &H8000000F
    gamemake.mak3a.BorderColor = &H8000000F: gamemake.mak3b.BorderColor = &H8000000F
End If
If gamemc >= 1000000 Then
    gamemake.gameke1(3).ForeColor = &HFF00&
    gamemake.mak4a.BorderColor = &HFF00&: gamemake.mak4b.BorderColor = &HFF00&
ElseIf gamemc < -100000000 Then
    gamemake.gameke1(3).ForeColor = &H8000000F
    gamemake.mak4a.BorderColor = &H8000000F: gamemake.mak4b.BorderColor = &H8000000F
End If
End Function



Public Function usewpin() '使用物品
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "药" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        hp = hp + gameyao_huifu
        gameyao = gameyao - 1
        gameyao_huifu = gameyao_huifu + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(1).Caption = "熟肉" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        eat = eat + 70
        gamesrou = gamesrou - 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(2).Caption = "error" And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then

    ElseIf gamemake.gameke1(3).Caption = "error" And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then

    End If
ElseIf wppage = 2 Then
End If
End Function



Public Function usewpinzs() '使用物品信息
Static xxnum As Integer
Static zxx As Variant
Static wancheng As Integer
If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '因为会保存 所以重置初始值
ElseIf wancheng = 0 Then
End If
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "药" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        zxx = "使用了一个" & gamemake.gameke1(0).Caption & "恢复了" & gameyao_huifu - 1 & "点生命值"
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '防止出现无信息问题
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' 刷新数据
        End If
    ElseIf gamemake.gameke1(1).Caption = "熟肉" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        zxx = "使用了一个" & gamemake.gameke1(1).Caption & "恢复了70饱食度"
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '防止出现无信息问题
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' 刷新数据
        End If
    ElseIf gamemake.gameke1(2).Caption = "error" & gamemy - 1 And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
        zxx = "制造了一个" & "麻衣LV" & gamemy
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '防止出现无信息问题
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' 刷新数据
        End If
    ElseIf gamemake.gameke1(3).Caption = "error" & gamemg - 1 And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
        zxx = "制造了一个" & "木棍LV" & gamemg
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '防止出现无信息问题
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' 刷新数据
        End If
    End If
End If
End Function



Public Function getzsxingxi() '未使用
gameget.getzs.Caption = "当前生命值:" & love & "/" & lovemax & "   当前饱食度:" & eat & "/" & eatmax _
& vbCrLf & "药类每次使用都会"
End Function '未使用
Public Function zdopen() '启动战斗状态awa
game.xzd.Enabled = True: game.xzd.Visible = True
game.xtp.Enabled = True: game.xtp.Visible = True
'开启战斗按钮
game.xqj.Enabled = False: game.xqj.Visible = False
game.xsave.Enabled = False: game.xsave.Visible = False
game.xts.Enabled = False: game.xts.Visible = False
game.xtx.Enabled = False: game.xtx.Visible = False
game.xzz.Enabled = False: game.xzz.Visible = False
game.xbb.Enabled = False: game.xbb.Visible = False
game.xsz.Enabled = False: game.xsz.Visible = False
game.xsj.Enabled = False: game.xsj.Visible = False
'关闭交互按钮

End Function



Public Function zdexit()
game.xzd.Enabled = False: game.xzd.Visible = False
game.xtp.Enabled = False: game.xtp.Enabled = False
'关闭战斗按钮
game.xqj.Enabled = True: game.xqj.Visible = True
game.xsave.Enabled = True: game.xsave.Visible = True
game.xts.Enabled = True: game.xts.Visible = True
game.xtx.Enabled = True: game.xtx.Visible = True
game.xzz.Enabled = True: game.xzz.Visible = True
game.xbb.Enabled = True: game.xbb.Visible = True
game.xsz.Enabled = True: game.xsz.Visible = True
game.xsj.Enabled = True: game.xsj.Visible = True
'开启交互按钮
End Function

