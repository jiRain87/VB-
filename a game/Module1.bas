Attribute VB_Name = "Module1"
Public m, z, zm, zz, yhm '�û�������
Public hp%, hpmax%, eat%, eatmax%, att%, def%, health%, gold%, phy%, phymax%, game_day As Integer, game_km As Integer '������Ϸ�û�����
'�ֱ��� ����ֵ �������� ��ʳ�� ��ʳ������ ���� ���� ���� ��� ���� �������� ��Ϸ���� ��Ϸ·��
Public ����hp, ����hpmax, ����att, ����def, ����phymax '�����������ֵ
Public gwhp, gwatt, gwdef '�����������
'�ֱ��ǹ��� Ѫ�� ���� ����
Public sjsj As Integer '�¼�
Public sjsjwp As Integer, sjsjwps1, sjsjwps2, sjsjwps3 As Integer '�����Ʒ
'��������Ʒ�¼� ��Ʒ����
Public sjsjgw As Integer '��������
Public sjp As Integer '�����ֽ
Public gameend As Integer '�˳�ʱ���ص���ֵ
Public gameym%, gameyc%, gamexyc%, gamemc%, gamerou As Integer '�����������ȫΪ����
'�����ҷֱ�Ϊ ���� ҩ�� ޹�²� ľ�� ����
Public gameapple%, gamesrou%, gameyao%, gameyan As Integer '������Ʒ
'������ ƻ�� ���� ҩ ��
Public gameyao_huifu As Integer '��ҩ�ָ���hp
Public gamemg%, gamemy As Integer '����װ��
'Ϊľ�� ����
Public gamemgs%, gamemys As Integer '����������������
Public wppage%, wpins As Integer '����ҳҳ�� ��Ʒ��λ��
Public zzbb As Integer '�ж������컹�Ǳ��� ����0 ����1
Public Function loadzz() '����浵
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
Put #1, 15, ����hp
Put #1, 16, ����hpmax
Put #1, 17, ����att
Put #1, 18, ����def
Put #1, 19, ����phymax
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



Public Function loadget() '��ȡ�浵
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
Get #1, 15, ����hp
Get #1, 16, ����hpmax
Get #1, 17, ����att
Get #1, 18, ����def
Get #1, 19, ����phymax
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



Public Function shuxing() 'δʹ��
att = gamemg * 10
lovemax = ����lovemax + gamemy * 10
End Function



Public Sub nagame() 'ˢ�½���\
If hp > hpmax Then
    hp = hpmax
End If
If phy > phymax Then
    phy = phymax
End If
game.gameday.Caption = "����:" & game_day
game.gamekm.Caption = "·��:" & game_km & "km"
game.gamegold.Caption = "���:" & gold
game.gamephy.Caption = "����:" & phy & "/" & phymax
game.gameatt.Caption = "����:" & att
game.gamedef.Caption = "����:" & def
game.gamehp.Caption = "����ֵ:" & hp & "/" & hpmax
game.gamehealth.Caption = "����:" & health
End Sub



Public Function newpage() 'ˢ������ �Լ���ҳ���µı䶯
If wppage = 1 Then
    gamemake.gameke1(0).Caption = "ҩ"
    gamemake.ke1.Caption = gameyc & "/2ҩ��" & vbCrLf & "���Իָ�" & gameyao_huifu & "��Ѫ������Ҫ�Ĳ���Ʒ" & vbCrLf & "��ǰӵ��" & gameyao
    gamemake.gameke1(1).Caption = "����"
    gamemake.ke2.Caption = gamerou & "/1 ��" & gamemc & "/2 ľ��" & vbCrLf & "���Իָ�70�㱥ʳ�� ʮ�ֺóԵ�Ұζ��" & vbCrLf & "��ǰӵ��" & gamesrou
    gamemake.gameke1(2).Caption = "����LV" & gamemy
    gamemake.ke3.Caption = gameym & "/" & gamemys & "����" & vbCrLf & "ÿ���ṩ" & gamemy * 10 & "������Ͱٷ�֮��ĸ�" _
    & vbCrLf & "��ǰ�ṩ" & gamemy * 5 & "�ĸ񵲼���"
    gamemake.gameke1(3).Caption = "ľ��LV" & gamemg
    gamemake.ke4.Caption = gamemc & "/" & gamemgs & "ľ��" & vbCrLf & "ÿ���ṩ" & gamemg * 10 & "�㹥���Ͱٷ�֮��ı���" _
    & vbCrLf & "��ǰ�ṩ" & gamemg * 5 & "�ı�������"

    gamemake.spage.Enabled = False: gamemake.spage.Visible = False
    gamemake.xpage.Enabled = True: gamemake.xpage.Visible = True
ElseIf wppage = 2 Then
    gamemake.gameke1(0).Caption = "δ����"
    gamemake.ke1.Caption = gameyc & "/2δ����" & Asc(13) & "δ����"
    gamemake.gameke1(1).Caption = "δ����"
    gamemake.ke2.Caption = gamerou & "/1 δ����" & gamemc & "/2 δ����" & Asc(13) & "δ����"
    gamemake.gameke1(2).Caption = "δ����" & gamemy
    gamemake.ke3.Caption = gameym & "/" & gamemys & Asc(13) & "δ����" _
    & Asc(13) & "δ����" & gamemy * 5 & "δ����"
    gamemake.gameke1(3).Caption = "δ����" & gamemg
    gamemake.ke4.Caption = gamemc & "/" & gamemgs & Asc(13) & "δ������ı���" _
    & Asc(13) & "��δ���ӹ�" & gamemg * 5 & "δ������"
    
    gamemake.spage.Enabled = True: gamemake.spage.Visible = True
    gamemake.xpage.Enabled = False: gamemake.xpage.Visible = False
End If
End Function



Public Function newpagegreen() 'ˢ������
If gameyc >= 2 Then
    gamemake.gameke1(0).ForeColor = &HFF00&
    gamemake.mak1a.BorderColor = &HFF00&: gamemake.mak1b.BorderColor = &HFF00& '����������� ����ɫŪ����ɫ
ElseIf gameyc < 2 Then
    gamemake.gameke1(0).ForeColor = &H8000000F
    gamemake.mak1a.BorderColor = &H8000000F: gamemake.mak1b.BorderColor = &H8000000F '������ֲ������� ����ɫŪ�ɰ�ɫ
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



Public Function makewpin() '������Ʒ
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "ҩ" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        gameyc = gameyc - 2
        gameyao = gameyao + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(1).Caption = "����" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        gamemc = gamemc - 2
        gamerou = gamerou - 1
        gamesrou = gamesrou + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(2).Caption = "����LV" & gamemy And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
        gameym = gameym - gamemys 'ȥ������
        gamemy = gamemy + 1 '���ӵȼ�
        hpmax = hpmax + 50
        hp = hp + 50 '������Ӧ�õ�
        gamemys = 5 + gamemy * 5 '���Ӵ���
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(3).Caption = "ľ��LV" & gamemg And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
        gamemc = gamemc - gamemgs
        gamemg = gamemg + 1
        att = att + 10
        gamemgs = 5 + gamemg * 5
        gamemake.maketime.Enabled = True
    End If
    
End If
End Function


Public Function makewpinzs() '��Ʒ������Ϣ
If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '��Ϊ�ᱣ�� �������ó�ʼֵ
ElseIf wancheng = 0 Then
End If
If wppage = 1 Then
    If gameke1(0).Caption = "ҩ" And gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        zxx = "������һ��" & gameke1(0).Caption
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
End Function

Public Function newpagebb() ' �����ķ�ҳ��Ϣˢ��
If wppage = 1 Then
    gamemake.gameke1(0).Caption = "ҩ"
    gamemake.ke1.Caption = "�ָ�" & gameyao_huifu & "������" & vbCrLf & "��ǰӵ��" & gameyao
    gamemake.gameke1(1).Caption = "����"
    gamemake.ke2.Caption = "���Իָ�70�㱥ʳ�� ʮ�ֺóԵ�Ұζ��" & vbCrLf & "��ǰӵ��" & gamesrou
    gamemake.gameke1(2).Caption = "δ��"
    gamemake.ke3.Caption = "��δ��"
    gamemake.gameke1(3).Caption = ��δ��
    gamemake.ke4.Caption = "��δ��"
    
    gamemake.spage.Enabled = False: gamemake.spage.Visible = False
    gamemake.xpage.Enabled = True: gamemake.xpage.Visible = True
ElseIf wppage = 2 Then
    gamemake.gameke1(0).Caption = "δ����"
    gamemake.ke1.Caption = "δ����" & Asc(13) & "δ����"
    gamemake.gameke1(1).Caption = "δ����"
    gamemake.ke2.Caption = "δ����/2 δ����" & Asc(13) & "δ����"
    gamemake.gameke1(2).Caption = "δ����"
    gamemake.ke3.Caption = "/" & Asc(13) & "δ����" _
    & Asc(13) & "δ����" & "δ����"
    gamemake.gameke1(3).Caption = "δ����"
    gamemake.ke4.Caption = "/" & Asc(13) & "δ���ӵı���" _
    & Asc(13) & "��δ���ӹ�δ������"
    
    gamemake.spage.Enabled = True: gamemake.spage.Visible = True
    gamemake.xpage.Enabled = False: gamemake.xpage.Visible = False
End If
End Function




Public Function newpagegreenbb() '������Ч��ˢ��
If gameyao > 0 Then
    gamemake.gameke1(0).ForeColor = &HFF00&
    gamemake.mak1a.BorderColor = &HFF00&: gamemake.mak1b.BorderColor = &HFF00& '����������� ����ɫŪ����ɫ
ElseIf gameyao <= 0 Then
    gamemake.gameke1(0).ForeColor = &H8000000F
    gamemake.mak1a.BorderColor = &H8000000F: gamemake.mak1b.BorderColor = &H8000000F '������ֲ������� ����ɫŪ�ɰ�ɫ
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



Public Function usewpin() 'ʹ����Ʒ
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "ҩ" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        hp = hp + gameyao_huifu
        gameyao = gameyao - 1
        gameyao_huifu = gameyao_huifu + 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(1).Caption = "����" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        eat = eat + 70
        gamesrou = gamesrou - 1
        gamemake.maketime.Enabled = True
    ElseIf gamemake.gameke1(2).Caption = "error" And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then

    ElseIf gamemake.gameke1(3).Caption = "error" And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then

    End If
ElseIf wppage = 2 Then
End If
End Function



Public Function usewpinzs() 'ʹ����Ʒ��Ϣ
Static xxnum As Integer
Static zxx As Variant
Static wancheng As Integer
If wancheng = 1 Then
        gamemake.makezs.Caption = ""
        xxnum = 0 '��Ϊ�ᱣ�� �������ó�ʼֵ
ElseIf wancheng = 0 Then
End If
If wppage = 1 Then
    If gamemake.gameke1(0).Caption = "ҩ" And gamemake.gameke1(0).ForeColor = &HFF00& And wpins = 0 Then
        zxx = "ʹ����һ��" & gamemake.gameke1(0).Caption & "�ָ���" & gameyao_huifu - 1 & "������ֵ"
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' ˢ������
        End If
    ElseIf gamemake.gameke1(1).Caption = "����" And gamemake.gameke1(1).ForeColor = &HFF00& And wpins = 1 Then
        zxx = "ʹ����һ��" & gamemake.gameke1(1).Caption & "�ָ���70��ʳ��"
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' ˢ������
        End If
    ElseIf gamemake.gameke1(2).Caption = "error" & gamemy - 1 And gamemake.gameke1(2).ForeColor = &HFF00& And wpins = 2 Then
        zxx = "������һ��" & "����LV" & gamemy
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' ˢ������
        End If
    ElseIf gamemake.gameke1(3).Caption = "error" & gamemg - 1 And gamemake.gameke1(3).ForeColor = &HFF00& And wpins = 3 Then
        zxx = "������һ��" & "ľ��LV" & gamemg
        If xxnum >= Len(zxx) Then
        xxnum = 0
        gamemake.makezs.Caption = ""
        End If '��ֹ��������Ϣ����
        xxnum = xxnum + 1
        gamemake.makezs.Caption = gamemake.makezs.Caption & Mid(zxx, xxnum, 1)
        wancheng = 0
        If gamemake.makezs.Caption = zxx Then
            gamemake.maketime.Enabled = False
            wancheng = 1
            Call newpagebb
            Call newpagegreenbb ' ˢ������
        End If
    End If
End If
End Function



Public Function getzsxingxi() 'δʹ��
gameget.getzs.Caption = "��ǰ����ֵ:" & love & "/" & lovemax & "   ��ǰ��ʳ��:" & eat & "/" & eatmax _
& vbCrLf & "ҩ��ÿ��ʹ�ö���"
End Function 'δʹ��
Public Function zdopen() '����ս��״̬awa
game.xzd.Enabled = True: game.xzd.Visible = True
game.xtp.Enabled = True: game.xtp.Visible = True
'����ս����ť
game.xqj.Enabled = False: game.xqj.Visible = False
game.xsave.Enabled = False: game.xsave.Visible = False
game.xts.Enabled = False: game.xts.Visible = False
game.xtx.Enabled = False: game.xtx.Visible = False
game.xzz.Enabled = False: game.xzz.Visible = False
game.xbb.Enabled = False: game.xbb.Visible = False
game.xsz.Enabled = False: game.xsz.Visible = False
game.xsj.Enabled = False: game.xsj.Visible = False
'�رս�����ť

End Function



Public Function zdexit()
game.xzd.Enabled = False: game.xzd.Visible = False
game.xtp.Enabled = False: game.xtp.Enabled = False
'�ر�ս����ť
game.xqj.Enabled = True: game.xqj.Visible = True
game.xsave.Enabled = True: game.xsave.Visible = True
game.xts.Enabled = True: game.xts.Visible = True
game.xtx.Enabled = True: game.xtx.Visible = True
game.xzz.Enabled = True: game.xzz.Visible = True
game.xbb.Enabled = True: game.xbb.Visible = True
game.xsz.Enabled = True: game.xsz.Visible = True
game.xsj.Enabled = True: game.xsj.Visible = True
'����������ť
End Function

