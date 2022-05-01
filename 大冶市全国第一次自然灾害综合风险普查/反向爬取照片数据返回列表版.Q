[General]
SyntaxVersion=2
BeginHotkey=38
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=37
StopHotkeyMod=8
RunOnce=0
EnableWindow=
MacroID=9b8719fb-b5c3-4123-93d0-cc88c6364db9
Description=反向爬取照片数据
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
For i=1 To 10
    //调查
    cr000 = GetPixelColor(1880,445)
    Do Until cr000 ="F9F9F9"  
        Delay 200
        cr000 = GetPixelColor(1880,445)  
    Loop
    MoveTo 1783,483+i*39
    cr00 = GetPixelColor(1896,483+i*39)
    Do Until cr00 ="ECECEC"  
        Delay 200
        cr00 = GetPixelColor(1896,483+i*39)  
    Loop
    Delay 200
    LeftClick 1
    //安全措施
    crchushi = GetPixelColor(1151, 283)
    Do Until crchushi ="FB803B"  
        Delay 200
        crchushi = GetPixelColor(1151, 283)  
    Loop
    Delay 200
    IfColor 1266,330, "AA9C01", 2 Then
        crzhu = GetPixelColor(1151, 735)
        Do Until crzhu ="FB803B"  
            Delay 200
            crzhu = GetPixelColor(1151, 735)  
        Loop
        Delay 200
        fwlx = "住宅"
        IfColor 1266, 847, "AA9C01", 2 Then
            dcqk = "不需要调查"
        Else 
            dcqk = "需要调查"
        End If
    Else 
        crzhu = GetPixelColor(1637,589)
        Do Until crzhu ="AA9C01"  
            Delay 200
            crzhu = GetPixelColor(1637,589)  
        Loop
        Delay 200
        IfColor 1266,690, "AA9C01", 2 Then
            dcqk = "不需要调查"
        Else 
            dcqk = "需要调查"
        End If
        fwlx = "非住宅"
    End If
    //复制主房编号
    MoveTo 1382, 280
    Delay 50
    LeftDoubleClick 1
    Delay 100
    Call 复制()
    zfbh = Plugin.Sys.GetCLB 
    Delay 100
    //跳表三
    MoveTo 1382, 220
    Delay 100
    LeftClick 1
    Delay 300
    If fwlx = "住宅" Then 
        IfColor 1286, 325, "AA9C01", 2 Then
            表三y = 833
        Else 
            表三y = 714
        End If
        cr1 = GetPixelColor(1221,表三y)
        cr2 = GetPixelColor(1411,表三y)
        cr3 = GetPixelColor(1601,表三y)
        //1
        MoveTo 1291,表三y-100
        Delay 100
        Call 照片另存()
        SayString zfbh&"(1)"
        Delay 100
        KeyPress "Enter", 1
        //2
        If cr2 = "8F7031" or cr3 = "8F7031" or cr2 = "8F7132" or cr3 = "8F7132" Then 
            Delay 500
            MoveTo 1479,表三y-100
            Delay 100
            Call 照片另存()
            SayString zfbh&"(2)"
            Delay 100
            KeyPress "Enter", 1
        End If
        //3
        If cr3 = "8F7031" or cr3 = "8F7132" Then 
            Delay 500
            MoveTo 1665,表三y-100
            Delay 100
            Call 照片另存()
            SayString zfbh&"(3)"
            Delay 100
            KeyPress "Enter", 1
        End If
    Else 
        IfColor 1286, 325, "AA9C01", 2 Then
            表三y = 848
        Else 
            表三y = 729
        End If
        cr1 = GetPixelColor(1221,表三y)
        cr2 = GetPixelColor(1411,表三y)
        cr3 = GetPixelColor(1601,表三y)
        //1
        MoveTo 1291,表三y-100
        Delay 100
        Call 照片另存()
        SayString zfbh&"(1)"
        Delay 100
        KeyPress "Enter", 1
        //2
        If cr2 = "8F7031" or cr3 = "8F7031" or cr2 = "8F7132" or cr3 = "8F7132" Then 
            Delay 500
            MoveTo 1479,表三y-100
            Delay 100
            Call 照片另存()
            SayString zfbh&"(2)"
            Delay 100
            KeyPress "Enter", 1
        End If
        //3
        If cr3 = "8F7031" or cr3 = "8F7132" Then 
            Delay 500
            MoveTo 1665,表三y-100
            Delay 100
            Call 照片另存()
            SayString zfbh&"(3)"
            Delay 100
            KeyPress "Enter", 1
        End If
    End If
    If fwlx = "住宅" Then 
        //检查辅助用房
        MoveTo 1151,220
        Delay 100
        LeftClick 1
        Delay 500
        MoveTo 1382,735
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 100
        fzfbh1 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh1 = zfbh Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh1
            Delay 100
            KeyPress "Enter", 1
        End If 
        //滚动
        MoveTo 1888,456
        Delay 100
        LeftClick 1
        Delay 100
        MouseWheel - 8 
        Delay 200
        //2
        MoveTo 1390,276
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh2 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh2 = fzfbh1 or Len(fzfbh2)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh2
            Delay 100
            KeyPress "Enter", 1 
        End If
        //3
        MoveTo 1388,620 
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh3 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh3 = fzfbh2 or Len(fzfbh3)<>15 Then 
            num = 2
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh3
            Delay 100
            KeyPress "Enter", 1
        End If
        //滚动
        MoveTo 1888,456
        Delay 100
        LeftClick 1
        Delay 100
        MouseWheel - 7
        Delay 200
        //4
        MoveTo 1388,266
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh4 = Plugin.Sys.GetCLB
        Delay 50
        fzfbhs4 = CStr(fzfbh4)
        Delay 50
        If fzfbh4 = fzfbh3 or Len(fzfbh4)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh4
            Delay 100
            KeyPress "Enter", 1
        End If
        //5
        MoveTo 1381,608
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh5 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh5 = fzfbh4 or Len(fzfbh5)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh5
            Delay 100
            KeyPress "Enter", 1
        End If
        //滚动
        MoveTo 1888,456
        Delay 100
        LeftClick 1
        Delay 100
        MouseWheel - 6
        Delay 200
        //6
        MoveTo 1381,352
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh6 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh6 = fzfbh5 or Len(fzfbh6)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh6
            Delay 100
            KeyPress "Enter", 1
        End If
        //7
        MoveTo 1381,697
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh7 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh7 = fzfbh6 or Len(fzfbh7)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh7
            Delay 100
            KeyPress "Enter", 1
        End If
        //滚动
        MoveTo 1888,456
        Delay 100
        LeftClick 1
        Delay 100
        MouseWheel - 7
        Delay 200
        //8
        MoveTo 1320,340
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh8 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbh8 = fzfbh7 or Len(fzfbh8)<>15 Then 
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh8
            Delay 100
            KeyPress "Enter", 1
        End If
        //9
        MoveTo 1381,683
        Delay 50
        LeftDoubleClick 1
        Delay 50
        Call 复制()
        Delay 50
        fzfbh9 = Plugin.Sys.GetCLB
        Delay 50
        If fzfbhs9 = fzfbhs8 or Len(fzfbhs9)<>15 Then
            Goto 完成
        Else 
            MoveR 0, 70
            Call 照片另存()
            SayString fzfbh9
            Delay 100
            KeyPress "Enter", 1
        End If
    End If
    Rem 完成
    //保存返回
    Delay 100
    MoveTo 1858,211
    Delay 100
    LeftClick 1
//    cr1 = GetPixelColor(1082,633)
//    Do While cr1 ="968A00"  
//        Delay 200
//        cr1 = GetPixelColor(1082,633)  
//    Loop
//    Delay 100
//    MoveTo 920,601
//    Delay 100
//    LeftClick 1
//    If dcqk = "不需要调查" Then 
//        Delay 1000
//        MoveTo 996,589
//        Delay 100
//        LeftClick 1
//    End If
    
Next
Delay 1000
MoveTo 1746, 935
Delay 200
LeftClick 1
Delay 1000
Sub 复制()
    KeyDown 17, 1
    KeyPress 67, 1
    KeyUp 17, 1
End Sub
Sub 辅助用房信息获取()
    //安全措施
    crfz = GetPixelColor(1210,515)
    Do Until crfz ="EEEEEE"  
        Delay 200
        crfz = GetPixelColor(1210,515)  
    Loop
    Delay 300
    cryt1 = GetPixelColor(702,557)
    cryt2 = GetPixelColor(765,557)
    cryt3 = GetPixelColor(828,557)
    cryt4 = GetPixelColor(891,557)
    cryt5 = GetPixelColor(954,557)
    cryt6 = GetPixelColor(702, 593)
    If cryt1 <> "FFFFFF" Then 
        yt = "厨房"
    ElseIf cryt2 <> "FFFFFF" Then
        yt = "厕所"
    ElseIf cryt3 <> "FFFFFF" Then
        yt = "车库"
    ElseIf cryt4 <> "FFFFFF" Then
        yt = "仓库"
    ElseIf cryt5 <> "FFFFFF" Then
        yt = "养殖圈舍"
    ElseIf cryt6 <> "FFFFFF" Then
        MoveTo 799,592
        Delay 50
        LeftClick 1
        Delay 50
        Call 全选()
        Delay 50
        Call 复制()
        Delay 50
        ytqt = Plugin.Sys.GetCLB
        Delay 50
        yt = "其他" & ytqt
    End If
    //表二 
    MoveTo 683,264
    Delay 50
    LeftClick 1
    Delay 400
    craqjd = GetPixelColor(711, 326)
    If craqjd = "FFFFFF" or craqjd = "FEFEFE"  or craqjd = "FEFFFE" or craqjd = "FEFFFF" or craqjd = "FFFFFE" Then 
        aqjd = "否"
    Else 
        aqjd = "是"
    End If
    //表三
    MoveTo 833,264
    Delay 50
    LeftClick 1
    Delay 500
    //审批手续
    crydsx = GetPixelColor(713,422)
    crjssx = GetPixelColor(1052,422)
    crsgsx = GetPixelColor(713,471)
    cryssx = GetPixelColor(1052,471)
    crspsx = GetPixelColor(713,520)
    crdjsx = GetPixelColor(1052,520)
    If crydsx = "FFFFFF" or crydsx = "FEFEFE"  or crydsx = "FEFFFE" or crydsx = "FEFFFF" or crydsx = "FFFFFE" Then 
        ydsx = "无"
    Else 
        ydsx = "有"
    End If
    If crjssx= "FFFFFF" or crjssx = "FEFEFE"  or crjssx = "FEFFFE" or crjssx = "FEFFFF"  or crjssx = "FFFFFE"  Then 
        jssx = "无"
    Else 
        jssx = "有"
    End If
    If crsgsx = "FFFFFF" or crsgsx = "FEFEFE"  or crsgsx = "FEFFFE" or crsgsx = "FEFFFF"  or crsgsx = "FFFFFE" Then 
        sgsx = "无"
    Else 
        sgsx = "有"
    End If
    If cryssx= "FFFFFF" or cryssx = "FEFEFE"  or cryssx = "FEFFFE" or cryssx = "FEFFFF"  or cryssx = "FFFFFE"  Then 
        yssx= "无"
    Else 
        yssx= "有"
    End If
    If crspsx = "FFFFFF" or crspsx = "FEFEFE"  or crspsx = "FEFFFE" or crspsx = "FEFFFF"  or crspsx = "FFFFFE" Then 
        spsx = "无"
    Else 
        spsx = "有"
    End If
    If crdjsx  = "FFFFFF" or crdjsx = "FEFEFE"  or crdjsx = "FEFFFE" or crdjsx = "FEFFFF"  or crdjsx = "FFFFFE" Then 
        djsx  = "无"
    Else 
        djsx  = "有"
    End If
    //用能情况
    cn1 = GetPixelColor(701,623)
    cn2 = GetPixelColor(764,623)
    cn3 = GetPixelColor(813,623)
    cn4 = GetPixelColor(890,623)
    cs1 = GetPixelColor(701,664)
    cs2 = GetPixelColor(764,664)
    cs3 = GetPixelColor(813,664)
    cs4 = GetPixelColor(890,664)
    //采暖其他
    cr341 = GetPixelColor(1175,610)
    cr342 = GetPixelColor(1175,630)
    cr343 = GetPixelColor(1190,615)
    //炊事其他
    cr351 = GetPixelColor(1175,650)
    cr352 = GetPixelColor(1173,671)
    cr353 = GetPixelColor(1190,661)
    If cn1 = "FFFFFF" or cn1 = "FEFEFE"  or cn1 = "FEFFFE" or cn1 = "FEFFFF" or cn1 = "FFFFFE" or cn1 = "FFFEFE"  or cn1 = "FEFEFF"   Then 
        cn1 = ""
    Else 
        cn1 = "煤炭"
    End If
    If cn2 = "FFFFFF" or cn2 = "FEFEFE"  or cn2 = "FEFFFE" or cn2 = "FEFFFF" or cn2 = "FFFFFE" or cn2 = "FFFEFE"  or cn2 = "FEFEFF"   Then 
        cn2 = ""
    Else 
        cn2 = "电"
    End If
    If cn3 = "FFFFFF" or cn3 = "FEFEFE"  or cn3 = "FEFFFE" or cn3 = "FEFFFF" or cn3 = "FFFFFE" or cn3 = "FFFEFE"   or cn3 = "FEFEFF"  Then 
        cn3 = ""
    Else 
        cn3 = "液化气"
    End If
    If cn4 = "FFFFFF" or cn4 = "FEFEFE"  or cn4 = "FEFFFE" or cn4 = "FEFFFF" or cn4 = "FFFFFE" or cn4 = "FFFEFE"   or cn4 = "FEFEFF"  Then 
        cn4 = ""
    Else 
        cn4 = "天然气"
    End If
    If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
        cn5 = ""
    Else 
        MoveTo 1060,618
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
    End If
    If cs1 = "FFFFFF" or cs1 = "FEFEFE"  or cs1 = "FEFFFE" or cs1 = "FEFFFF"  or cs1 = "FFFFFE" or cs1 = "FFFEFE"  or cs1 = "FEFEFF"  Then 
        cs1 = ""
    Else 
        cs1 = "煤炭"
    End If
    If cs2 = "FFFFFF" or cs2 = "FEFEFE"  or cs2 = "FEFFFE" or cs2 = "FEFFFF" or  cs2 = "FFFFFE" or cs2 = "FFFEFE" or cs2 = "FEFEFF"  Then 
        cs2 = ""
    Else 
        cs2 = "电"
    End If
    If cs3 = "FFFFFF" or cs3 = "FEFEFE"  or cs3 = "FEFFFE" or cs3 = "FEFFFF"  or cs3 = "FFFFFE" or cs3 = "FFFEFE"  or cs3 = "FEFEFF"  Then 
        cs3 = ""
    Else 
        cs3 = "液化气"
    End If
    If cs4 = "FFFFFF" or cs4 = "FEFEFE"  or cs4 = "FEFFFE" or cs4 = "FEFFFF"  or cs4 = "FFFFFE" or cs4 = "FFFEFE"  or cs4 = "FEFEFF"  Then 
        cs4 = ""
    Else 
        cs4 = "天然气"
    End If
    If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
        cs5 = ""
    Else 
        MoveTo 1060,660
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
    End If
    cnyn = cn1&cn2&cn3&cn4&cn5
    csyn = cs1 & cs2 & cs3 & cs4 & cs5
    //是否改造
    crsfgz = GetPixelColor(701, 753)
    If crsfgz = "FFFFFF" or crsfgz = "FEFEFE"  or crsfgz = "FEFFFE" or crsfgz = "FEFFFF" or  crsfgz = "FFFFFE" Then 
        sfgz = "是"
    Else 
        sfgz = "否"
    End If
    //取消
    MoveTo 995,872
    Delay 50
    LeftClick 1
    Delay 500
End Sub
Sub 全选()
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
End Sub
Sub 辅助用房信息填写()
    KeyPress "Right", 1
    Delay 30
    SayString yt
    KeyPress "Right", 1
    Delay 30
    SayString aqjd
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
    SayString sfgz
    Delay 30
    KeyPress "Left", 11
    Delay 50
    KeyPress "Enter", 1
    Delay 50
End Sub
Sub 照片另存()
    Delay 100
    RightClick 1
    //照片另存
    MoveR 70, 37
    Delay 100
    LeftClick 1
    crz= GetPixelColor(536,91)
    Do Until crz ="F7F6F5"  
        Delay 200
        crz = GetPixelColor(536,91)  
    Loop
    Delay 200
End Sub
