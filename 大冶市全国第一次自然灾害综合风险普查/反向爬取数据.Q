[General]
SyntaxVersion=2
BeginHotkey=38
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=45
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=2e8e8491-d112-4337-826b-cc8adcd37bd9
Description=反向爬取数据
Enable=1
AutoRun=0
[Repeat]
Type=1
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
//安全设置
cr1 = GetPixelColor(1151,283)
Do Until cr1 ="FB803B"  
    Delay 200
    cr1 = GetPixelColor(1151,283)  
Loop
Delay 500
IfColor 1266, 330, "FFFFFF", 0 Then
    //非住宅
    cr2 = GetPixelColor(1637,589)
    Do Until cr2 ="AA9C01"  
        Delay 200
        cr2 = GetPixelColor(1637,589)  
    Loop
Else 
    //住宅
    cr2 = GetPixelColor(1151,736)
    Do Until cr2 ="FB803B"  
        Delay 200
        cr2 = GetPixelColor(1151,736)  
    Loop
End If
//复制编号
MoveTo 1330,280
Delay 50
LeftDoubleClick 1
Delay 50
Call 复制()
Delay 50
fwbh = Plugin.Sys.GetCLB 
fwbh = CStr(fwbh)
//房屋类别
IfColor 1266, 330, "FFFFFF", 0 Then
    Delay 200
    fwlb = "非住宅"
    //现场调查情况
    IfColor 1266,690, "AA9C01", 0 Then
        dc = "不需要调查"
        Call 不需要调查()
        Goto 调查
    Else
        dc = "需要调查"
    End If
    Call 非住宅()
Else 
    Delay 200
    fwlb = "住宅"
    //现场调查情况
    IfColor 1266, 847, "AA9C01", 0 Then
        dc = "不需要调查"
        Call 不需要调查()
        Goto 调查
    Else
        dc = "需要调查"
    End If
    Call 住宅()
End If
Call 完成()
cr000 = GetPixelColor(1845,361)
Do Until cr000 ="F9F9F9"  
    Delay 200
    cr000 = GetPixelColor(1845,361)  
Loop
Call 表格写入()
//调查
Rem 调查
y = 405
MoveTo 1815,y
cr00 = GetPixelColor(1890,y+20)
Do Until cr00 ="ECECEC"  
    Delay 200
    cr00 = GetPixelColor(1890,y+20)  
Loop
Delay 200
LeftClick 1
Sub 住宅()
    //复制姓名
    MoveTo 1283,582
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    xm = Plugin.Sys.GetCLB
    Delay 50
    //复制身份证
    MoveTo 1282,631
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    sfz = Plugin.Sys.GetCLB
    //获取地址
    MoveTo 1277,511
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    zu = Plugin.Sys.GetCLB
    //组
    MoveTo 1446,513
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    lu = Plugin.Sys.GetCLB
    //跳转表二
    MoveTo 1256,221
    Delay 100
    LeftClick 1
    Delay 300
    //建筑层数，建筑面积
    MoveTo 1277,281
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    jzcs = Plugin.Sys.GetCLB
    MoveTo 1283,329
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    jzmj = Plugin.Sys.GetCLB
    //年份
    crn1 = GetPixelColor(1266,390)
    crn2 = GetPixelColor(1388,390)
    crn3 = GetPixelColor(1504,390)
    crn4 = GetPixelColor(1266,419)
    crn5 = GetPixelColor(1382,419)
    crn6 = GetPixelColor(1497,419)
    If crn1 <> "FFFFFF" Then 
        nf = "1980年及以前"
    ElseIf crn2 <> "FFFFFF" Then
        nf = "1980-1990年"
    ElseIf crn3 <> "FFFFFF" Then
        nf = "1991-2000年"
    ElseIf crn4 <> "FFFFFF" Then
        nf = "2001-2010年"
    ElseIf crn5 <> "FFFFFF" Then
        nf = "2010-2015年"
    ElseIf crn6 <> "FFFFFF" Then
        nf = "2016年及以后"
    Else 
        nf = "未填写"
    End If
    //结构类型 
    crjg1 = GetPixelColor(1266,463)
    crjg2 = GetPixelColor(1357,463)
    crjg3 = GetPixelColor(1448,463)
    crjg4 = GetPixelColor(1539,463)
    crjg5 = GetPixelColor(1266,492)
    crjg6 = GetPixelColor(1399,492)
    crjg7 = GetPixelColor(1476,492)
    If crjg1 <> "FFFFFF" Then 
        IfColor 1266,628,"FFFFFF",0 Then
            zsjg = "木或轻钢楼屋盖"
        Else    
            zsjg = "预制板"
        End If
        jglx = "砖石结构-"&zsjg
    ElseIf crjg2 <> "FFFFFF" Then
        IfColor 1266,580,"FFFFFF",0 Then
            tmjg = "木（竹）结构"
        Else    
            tmjg = "生土结构"
        End If
        jglx = "土木结构-"&tmjg
    ElseIf crjg3 <> "FFFFFF" Then
        jglx = "混杂结构"
    ElseIf crjg4 <> "FFFFFF" Then
        jglx = "窑洞"
    ElseIf crjg5 <> "FFFFFF" Then
        jglx = "钢筋混凝土结构"
    ElseIf crjg6 <> "FFFFFF" Then
        jglx = "钢结构"
    ElseIf crjg7 <> "FFFFFF" Then
        MoveTo 1325,525
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qtjg = Plugin.Sys.GetCLB
        Delay 50
        jglx = "其他结构-"&qtjg
    End If
    //建造方式
    If crjg1 <> "FFFFFF" Then 
        crjzfs1 = GetPixelColor(1266,725)
        crjzfs2 = GetPixelColor(1357,725)
        crjzfs3 = GetPixelColor(1476,725)
        crjzfs4 = GetPixelColor(1266, 762)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,814,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    ElseIf crjg2 <> "FFFFFF" Then
        crjzfs1 = GetPixelColor(1266,624)
        crjzfs2 = GetPixelColor(1357,624)
        crjzfs3 = GetPixelColor(1476,624)
        crjzfs4 = GetPixelColor(1266,661)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,713,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    Else 
        crjzfs1 = GetPixelColor(1266,580)
        crjzfs2 = GetPixelColor(1357,580)
        crjzfs3 = GetPixelColor(1476,580)
        crjzfs4 = GetPixelColor(1266,617)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,669,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    End If
    //表三
    MoveTo 1374,222
    Delay 100
    LeftClick 1
    Delay 300
    //判断
    IfColor 1339, 281, "FFFFFF", 1 Then
        zysj = "否"
    Else 
        zysj = "是"
    End If
    IfColor 1339, 325, "FFFFFF", 1 Then
        kzgz = "否"
        IfColor 1339,385 ,"FFFFFF", 1 Then
            kzjg = "否"
        Else 
            kzjg = "是"
        End If
        IfColor 1339,445, "FFFFFF", 1 Then
            ywlf = "否"
        Else 
            ywlf = "是"
        End If
    Else 
        kzgz = "是"
        IfColor 1339,504, "FFFFFF", 1 Then
            kzjg = "否"
        Else 
            kzjg = "是"
        End If
        IfColor 1339,564, "FFFFFF", 1 Then
            ywlf = "否"
        Else 
            ywlf = "是"
        End If
    End If
    //跳表四
    MoveTo 1519, 224
    Delay 50
    LeftClick 1
    Delay 100
    //条件判断
    cr表四 = GetPixelColor(1266,279)
    Do Until cr表四 = "AA9C01"  
        Delay 200
        cr表四 = GetPixelColor(1266,279)  
    Loop
    Delay 200
    crydsx = GetPixelColor(1278,378)
    crjssx = GetPixelColor(1604,378)
    crsgsx = GetPixelColor(1278,427)
    cryssx = GetPixelColor(1604,427)
    crspsx = GetPixelColor(1278,476)
    crdjsx = GetPixelColor(1604,476)
    If crydsx = "FFFFFF" Then 
        ydsx = "无"
    Else 
        ydsx = "有"
    End If
    If crjssx= "FFFFFF" Then 
        jssx = "无"
    Else 
        jssx = "有"
    End If
    If crsgsx = "FFFFFF" Then 
        sgsx = "无"
    Else 
        sgsx = "有"
    End If
    If cryssx= "FFFFFF" Then 
        yssx= "无"
    Else 
        yssx= "有"
    End If
    If crspsx = "FFFFFF" Then 
        spsx = "无"
    Else 
        spsx = "有"
    End If
    If crdjsx  = "FFFFFF" Then 
        djsx  = "无"
    Else 
        djsx  = "有"
    End If
    //用能情况
    //采暖用能
    cn1 = GetPixelColor(1266,580)
    cn2 = GetPixelColor(1329,580)
    cn3 = GetPixelColor(1378,580)
    cn4 = GetPixelColor(1455,580)
    cn5 = GetPixelColor(1532,580)
    cs1 = GetPixelColor(1266, 621)
    cs2 = GetPixelColor(1329, 621)
    cs3 = GetPixelColor(1378, 621)
    cs4 = GetPixelColor(1455, 621)
    cs5 = GetPixelColor(1532, 621)
    If cn1 <> "FFFFFF" Then 
        cn1 = "煤炭"
    Else 
        cn1 = ""
    End If
    If cn2 <> "FFFFFF" Then 
        cn2 = "电"
    Else 
        cn2 = ""
    End If
    If cn3 <> "FFFFFF" Then 
        cn3 = "液化气"
    Else 
        cn3 = ""
    End If
    If cn4 <> "FFFFFF" Then 
        cn4 = "天然气"
    Else 
        cn4 = ""
    End If
    If cn5 <> "FFFFFF" Then 
        MoveTo 1630,576
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qt = Plugin.Sys.GetCLB
        Delay 50
        cn5 = "其他"&qt
    Else 
        cn5 = ""
    End If
    If cs1 <> "FFFFFF" Then 
        cs1 = "煤炭"
    Else 
        cs1 = ""
    End If
    If cs2 <> "FFFFFF" Then 
        cs2 = "电"
    Else 
        cs2 = ""
    End If
    If cs3 <> "FFFFFF" Then 
        cs3 = "液化气"
    Else 
        cs3 = ""
    End If
    If cs4 <> "FFFFFF" Then 
        cs4 = "天然气"
    Else 
        cs4 = ""
    End If
    If cs5 <> "FFFFFF" Then 
        MoveTo 1630,617
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qt = Plugin.Sys.GetCLB
        Delay 50
        cs5 = "其他"&qt
    Else 
        cs5 = ""
    End If
    cnyn = cn1&cn2&cn3&cn4&cn4
    csyn = cs1 & cs2 & cs3 & cs4 & cs5
    //地址灾害隐患判断
    IfColor 1274, 808, "FFFFFF", 1 Then
        dzzhyh = "不存在"
        IfColor 1274, 857, "FFFFFF", 1 Then
            aqyhpd = "基本安全"
        Else 
            aqyhpd = "存在风险"
        End If
    Else 
        dzzhyh = "存在"
    End If
End Sub
Sub 非住宅()
    //复制姓名
    MoveTo 1276,538
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    xm = Plugin.Sys.GetCLB
    Delay 50
    //复制身份证
    MoveTo 1274,588 
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    sfz = Plugin.Sys.GetCLB
    //获取地址
    MoveTo 1276,465
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    zu = Plugin.Sys.GetCLB
    //组
    MoveTo 1445,469
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    lu = Plugin.Sys.GetCLB
    //跳转表二
    MoveTo 1256,221
    Delay 100
    LeftClick 1
    Delay 300
    //建筑层数，建筑面积
    MoveTo 1277,281
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    jzcs = Plugin.Sys.GetCLB
    MoveTo 1283,329
    Delay 30
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    Call 复制()
    Delay 50
    jzmj = Plugin.Sys.GetCLB
    //年份
    crn1 = GetPixelColor(1266,390)
    crn2 = GetPixelColor(1388,390)
    crn3 = GetPixelColor(1504,390)
    crn4 = GetPixelColor(1266,419)
    crn5 = GetPixelColor(1382,419)
    crn6 = GetPixelColor(1497,419)
    If crn1 <> "FFFFFF" Then 
        nf = "1980年及以前"
    ElseIf crn2 <> "FFFFFF" Then
        nf = "1980-1990年"
    ElseIf crn3 <> "FFFFFF" Then
        nf = "1991-2000年"
    ElseIf crn4 <> "FFFFFF" Then
        nf = "2001-2010年"
    ElseIf crn5 <> "FFFFFF" Then
        nf = "2010-2015年"
    ElseIf crn6 <> "FFFFFF" Then
        nf = "2016年及以后"
    Else 
        nf = "未填写"
    End If
    //结构类型 
    crjg1 = GetPixelColor(1266,463)
    crjg2 = GetPixelColor(1357,463)
    crjg3 = GetPixelColor(1448,463)
    crjg4 = GetPixelColor(1266,492)
    crjg5 = GetPixelColor(1329,492)
    crjg6 = GetPixelColor(1462,492)
    crjg7 = GetPixelColor(1266,529)
    If crjg1 <> "FFFFFF" Then 
        IfColor 1266,628,"FFFFFF",0 Then
            zsjg = "木或轻钢楼屋盖"
        Else    
            zsjg = "预制板"
        End If
        jglx = "砖石结构-"&zsjg
    ElseIf crjg2 <> "FFFFFF" Then
        IfColor 1266,580,"FFFFFF",0 Then
            tmjg = "木（竹）结构"
        Else    
            tmjg = "生土结构"
        End If
        jglx = "土木结构-"&tmjg
    ElseIf crjg3 <> "FFFFFF" Then
        jglx = "混杂结构"
    ElseIf crjg4 <> "FFFFFF" Then
        jglx = "窑洞"
    ElseIf crjg5 <> "FFFFFF" Then
        jglx = "钢筋混凝土结构"
    ElseIf crjg6 <> "FFFFFF" Then
        jglx = "钢结构"
    ElseIf crjg7 <> "FFFFFF" Then
        MoveTo 1345,527
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qtjg = Plugin.Sys.GetCLB
        Delay 50
        jglx = "其他结构-"&qtjg
    End If
    //建造方式
    If crjg1 <> "FFFFFF" Then 
        crjzfs1 = GetPixelColor(1266,725)
        crjzfs2 = GetPixelColor(1357,725)
        crjzfs3 = GetPixelColor(1476,725)
        crjzfs4 = GetPixelColor(1651,725)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,827,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    ElseIf crjg2 <> "FFFFFF" Then
        crjzfs1 = GetPixelColor(1266,624)
        crjzfs2 = GetPixelColor(1357,624)
        crjzfs3 = GetPixelColor(1476,624)
        crjzfs4 = GetPixelColor(1266,661)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,713,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    Else 
        crjzfs1 = GetPixelColor(1266,580)
        crjzfs2 = GetPixelColor(1357,580)
        crjzfs3 = GetPixelColor(1476,580)
        crjzfs4 = GetPixelColor(1266,617)
        If crjzfs1 = "AA9C01" Then 
            jzfs = "自行建造"
        ElseIf crjzfs2 = "AA9C01" Then
            jzfs = "建筑工匠建造"
        ElseIf crjzfs3 = "AA9C01" Then
            jzfs = "有资质的施工建造"
        ElseIf crjzfs4 = "AA9C01" Then
            jzfs = "其他建造"
        End If
        IfColor 1275,669,"FFFFFF",0 Then
            aqjd = "否"
        Else    
            aqjd = "是"
        End If
    End If
    //表三
    MoveTo 1374,222
    Delay 100
    LeftClick 1
    Delay 300
    //判断
    IfColor 1339, 281, "FFFFFF", 1 Then
        zysj = "否"
    Else 
        zysj = "是"
    End If
    IfColor 1339, 325, "FFFFFF", 1 Then
        kzgz = "否"
        IfColor 1339,385 ,"FFFFFF", 1 Then
            kzjg = "否"
        Else 
            kzjg = "是"
        End If
        IfColor 1339,445, "FFFFFF", 1 Then
            ywlf = "否"
        Else 
            ywlf = "是"
        End If
    Else 
        kzgz = "是"
        IfColor 1339,504, "FFFFFF", 1 Then
            kzjg = "否"
        Else 
            kzjg = "是"
        End If
        IfColor 1339,564, "FFFFFF", 1 Then
            ywlf = "否"
        Else 
            ywlf = "是"
        End If
    End If
    //跳表四
    MoveTo 1519, 224
    Delay 50
    LeftClick 1
    Delay 100
    //条件判断
    cr表四 = GetPixelColor(1266,279)
    Do Until cr表四 = "AA9C01"  
        Delay 200
        cr表四 = GetPixelColor(1266,279)  
    Loop
    Delay 200
    crydsx = GetPixelColor(1278,378)
    crjssx = GetPixelColor(1604,378)
    crsgsx = GetPixelColor(1278,427)
    cryssx = GetPixelColor(1604,427)
    crspsx = GetPixelColor(1278,476)
    crdjsx = GetPixelColor(1604,476)
    If crydsx = "FFFFFF" Then 
        ydsx = "无"
    Else 
        ydsx = "有"
    End If
    If crjssx= "FFFFFF" Then 
        jssx = "无"
    Else 
        jssx = "有"
    End If
    If crsgsx = "FFFFFF" Then 
        sgsx = "无"
    Else 
        sgsx = "有"
    End If
    If cryssx= "FFFFFF" Then 
        yssx= "无"
    Else 
        yssx= "有"
    End If
    If crspsx = "FFFFFF" Then 
        spsx = "无"
    Else 
        spsx = "有"
    End If
    If crdjsx  = "FFFFFF" Then 
        djsx  = "无"
    Else 
        djsx  = "有"
    End If
    //用能情况
    //采暖用能
    cn1 = GetPixelColor(1266,580)
    cn2 = GetPixelColor(1329,580)
    cn3 = GetPixelColor(1378,580)
    cn4 = GetPixelColor(1455,580)
    cn5 = GetPixelColor(1532,580)
    cs1 = GetPixelColor(1266, 621)
    cs2 = GetPixelColor(1329, 621)
    cs3 = GetPixelColor(1378, 621)
    cs4 = GetPixelColor(1455, 621)
    cs5 = GetPixelColor(1532, 621)
    If cn1 <> "FFFFFF" Then 
        cn1 = "煤炭"
    Else 
        cn1 = ""
    End If
    If cn2 <> "FFFFFF" Then 
        cn2 = "电"
    Else 
        cn2 = ""
    End If
    If cn3 <> "FFFFFF" Then 
        cn3 = "液化气"
    Else 
        cn3 = ""
    End If
    If cn4 <> "FFFFFF" Then 
        cn4 = "天然气"
    Else 
        cn4 = ""
    End If
    If cn5 <> "FFFFFF" Then 
        MoveTo 1630,576
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qt = Plugin.Sys.GetCLB
        Delay 50
        cn5 = "其他"&qt
    Else 
        cn5 = ""
    End If
    If cs1 <> "FFFFFF" Then 
        cs1 = "煤炭"
    Else 
        cs1 = ""
    End If
    If cs2 <> "FFFFFF" Then 
        cs2 = "电"
    Else 
        cs2 = ""
    End If
    If cs3 <> "FFFFFF" Then 
        cs3 = "液化气"
    Else 
        cs3 = ""
    End If
    If cs4 <> "FFFFFF" Then 
        cs4 = "天然气"
    Else 
        cs4 = ""
    End If
    If cs5 <> "FFFFFF" Then 
        MoveTo 1630,617
        Delay 30
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        qt = Plugin.Sys.GetCLB
        Delay 50
        cs5 = "其他"&qt
    Else 
        cs5 = ""
    End If
    cnyn = cn1&cn2&cn3&cn4&cn4
    csyn = cs1 & cs2 & cs3 & cs4 & cs5
    //地址灾害隐患判断
    IfColor 1274, 808, "FFFFFF", 1 Then
        dzzhyh = "不存在"
        IfColor 1274, 857, "FFFFFF", 1 Then
            aqyhpd = "基本安全"
        Else 
            aqyhpd = "存在风险"
        End If
    Else 
        dzzhyh = "存在"
    End If
End Sub
Sub 完成()
    //保存
    //    Delay 100
    //    MoveTo 1394,945
    //    Delay 100
    //    LeftClick 1
    //完成
    Delay 100
    MoveTo 1564,945
    Delay 100
    LeftClick 1
    cr1 = GetPixelColor(1082,633)
    Do While cr1 ="968A00"  
        Delay 200
        cr1 = GetPixelColor(1082,633)  
    Loop
    MoveTo 920, 632
    Delay 100
    LeftClick 1
End Sub
Sub 复制()
    KeyDown 17, 1
    KeyPress 67, 1
    KeyUp 17, 1
End Sub
Sub 全选()
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
End Sub
Sub 表格写入()
    //写入表格
    MoveTo 175,1059
    Delay 30
    LeftClick 1
    Delay 200
    SayString fwbh
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString dc
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString fwlb
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString xm
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString sfz
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString zu
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString lu
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString jzcs
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString jzmj
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString nf
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString jglx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString jzfs
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString aqjd
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString zysj
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString kzgz
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString kzjg
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString ywlf
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString ydsx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString jssx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString sgsx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString yssx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString spsx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString djsx
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString cnyn
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString csyn
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString dzzhyh
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString aqyhpd
    Delay 30
    KeyPress "Left", 26
    Delay 30
    KeyPress "Down", 1
    Delay 30
    MoveTo 123,1060
    Delay 30
    LeftClick 1
    Delay 100
End Sub
Sub 不需要调查()
    Call 完成()
    cr000 = GetPixelColor(1845,361)
    Do Until cr000 ="F9F9F9"  
        Delay 200
        cr000 = GetPixelColor(1845,361)  
    Loop
    //写入表格
    MoveTo 175,1059
    Delay 30
    LeftClick 1
    Delay 200
    SayString fwbh
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString dc
    Delay 30
    KeyPress "Right", 1
    Delay 30
    SayString fwlb
    Delay 30
    KeyPress "Left", 2
    Delay 30
    KeyPress "Down", 1
    Delay 30
    MoveTo 123,1060
    Delay 30
    LeftClick 1
    Delay 100
End Sub
//510681197902084120