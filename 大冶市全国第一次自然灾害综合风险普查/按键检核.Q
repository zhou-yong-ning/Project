[General]
SyntaxVersion=2
BeginHotkey=45
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=37
StopHotkeyMod=8
RunOnce=0
EnableWindow=
MacroID=6738976d-8045-484d-934a-0a605be99ce0
Description=按键检核
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
IfColor 1333,330, "AA9C01", 0 Then
    //非住宅
    cr2 = GetPixelColor(1758,412)
    Do Until cr2 ="343434"  
        Delay 200
        cr2 = GetPixelColor(1758,412)
    Loop
    Delay 200
    MoveTo 1744,414
    Delay 100
    LeftClick 1
    crsyr = GetPixelColor(1637,589)
    Do While crsyr ="AA9C01"  
        Delay 200
        crsyr = GetPixelColor(1637,589)  
    Loop
    Delay 200
    //选择村
    MoveTo 1741,645
    Delay 100
    LeftClick 1
Else 
    //住宅
    cr2 = GetPixelColor(1151,736)
    Do Until cr2 ="FB803B"  
        Delay 200
        cr2 = GetPixelColor(1151,736)  
    Loop
    Delay 200
    MoveTo 1731,457
    Delay 100
    LeftClick 1
    crsyr = GetPixelColor(1645,584)
    Do While crsyr ="AA9C01"  
        Delay 200
        crsyr = GetPixelColor(1645,584)  
    Loop
    Delay 200
    //选择村
    MoveTo 1707,690
    Delay 100
    LeftClick 1
End If
控制 = WaitKey()
Do While True
    If 控制 = 96 Then 
        Goto fanye
    End If
    If 控制 =39 Then 
        //		方向右键,点击通过
        Call 通过()
    ElseIf 控制 = 97 Then
        //    1
        Call 辅助用房()
    ElseIf 控制 = 98 Then
        //        2
        Call 全选()
        SayString "杨家洲生产队"
        Delay 200
        MoveTo 1912, 581
        Delay 50
        LeftClick 1
    ElseIf 控制 = 99 Then
        Call 全选()
        //        3
        SayString "510681197902084120"
        Delay 200
        MoveTo 1912, 581
        Delay 50
        LeftClick 1
    ElseIf 控制 = 100 Then
        //        4
        Call 存在风险()
    ElseIf 控制 = 101 Then
        //        5
        Call 基本安全()
    ElseIf 控制 = 102 Then
        //        6
        Call 查看照片()
    ElseIf 控制 = 103 Then
        //        7
        Call 住宅类()
        MoveTo 1896, 522
        Delay 100
        LeftClick 1
    ElseIf 控制 = 104 Then
        //        8
        Call 厕所类()
        MoveTo 1896, 522
        Delay 100
        LeftClick 1
    ElseIf 控制 = 105 Then
        Call 全选()
        //        9
        SayString "京南村村民委员会"
        MoveTo 1896, 522
        Delay 100
        LeftClick 1
    End If
    控制 = WaitKey()
    Rem fanye
    If 控制 = 96 Then
        Call 翻页()
    End If
Loop
Sub 全选()
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
End Sub
Sub 厕所类()
    //跳表四
    MoveTo 1546,221
    Delay 100
    LeftClick 1
    Delay 200
    crydsx = GetPixelColor(1278,378)
    If crydsx = "FFFFFF" or crydsx = "FEFEFE"  or crydsx = "FEFFFE" or crydsx = "FEFFFF" or crydsx = "FFFFFE" Then 
    Else 
        MoveTo 1424,373
        Delay 100
        LeftClick 1
        Delay 100
        MoveTo 1413,439
        Delay 100
        LeftClick 1
    End If
    //用能情况
    cn1 = GetPixelColor(1266,580)
    cn2 = GetPixelColor(1329,580)
    cn3 = GetPixelColor(1378,580)
    cn4 = GetPixelColor(1455,580)
    cs1 = GetPixelColor(1266,621)
    cs2 = GetPixelColor(1329,621)
    cs3 = GetPixelColor(1378,621)
    cs4 = GetPixelColor(1455,621)
    //采暖其他
    cr341 = GetPixelColor(1760,588)
    cr342 = GetPixelColor(1759,566)
    cr343 = GetPixelColor(1732,577)
    //炊事其他
    cr351 = GetPixelColor(1757,627)
    cr352 = GetPixelColor(1757,608)
    cr353 = GetPixelColor(1732,617)
    If cn1 = "FFFFFF" or cn1 = "FEFEFE"  or cn1 = "FEFFFE" or cn1 = "FEFFFF" or cn1 = "FFFFFE" or cn1 = "FFFEFE"  or cn1 = "FEFEFF"   Then 
    Else 
        MoveTo 1266,577
        Delay 100
        LeftClick 1
    End If
    If cn2 = "FFFFFF" or cn2 = "FEFEFE"  or cn2 = "FEFFFE" or cn2 = "FEFFFF" or cn2 = "FFFFFE" or cn2 = "FFFEFE"  or cn2 = "FEFEFF"   Then 
        //    MoveTo 1329,577
        //    Delay 100
        //    LeftClick 1
    Else 
        MoveTo 1329,577
        Delay 100
        LeftClick 1
    End If
    If cn3 = "FFFFFF" or cn3 = "FEFEFE"  or cn3 = "FEFFFE" or cn3 = "FEFFFF" or cn3 = "FFFFFE" or cn3 = "FFFEFE"   or cn3 = "FEFEFF"  Then 
    Else 
        MoveTo 1378,577
        Delay 100
        LeftClick 1
    End If
    If cn4 = "FFFFFF" or cn4 = "FEFEFE"  or cn4 = "FEFFFE" or cn4 = "FEFFFF" or cn4 = "FFFFFE" or cn4 = "FFFEFE"   or cn4 = "FEFEFF"  Then 
    Else 
        MoveTo 1455,577
        Delay 100
        LeftClick 1
    End If
    If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
        MoveTo 1533,577
        Delay 100
        LeftClick 1
        MoveTo 1633,577
        Delay 100
        LeftClick 1
        Delay 100
        SayString "无"
    Else 
        //    MoveTo 1533,577
        //    Delay 100
        //    LeftClick 1
    End If
    If cs1 = "FFFFFF" or cs1 = "FEFEFE"  or cs1 = "FEFFFE" or cs1 = "FEFFFF"  or cs1 = "FFFFFE" or cs1 = "FFFEFE"  or cs1 = "FEFEFF"  Then 
    Else 
        MoveTo 1266,618
        Delay 100
        LeftClick 1
    End If
    If cs2 = "FFFFFF" or cs2 = "FEFEFE"  or cs2 = "FEFFFE" or cs2 = "FEFFFF" or  cs2 = "FFFFFE" or cs2 = "FFFEFE" or cs2 = "FEFEFF"  Then 
        //    MoveTo 1329,618
        //    Delay 100
        //    LeftClick 1
    Else 
        MoveTo 1329,618
        Delay 100
        LeftClick 1
    End If
    If cs3 = "FFFFFF" or cs3 = "FEFEFE"  or cs3 = "FEFFFE" or cs3 = "FEFFFF"  or cs3 = "FFFFFE" or cs3 = "FFFEFE"  or cs3 = "FEFEFF"  Then 
        //    MoveTo 1378,618
        //    Delay 100
        //    LeftClick 1
    Else 
        MoveTo 1378,618
        Delay 100
        LeftClick 1
    End If
    If cs4 = "FFFFFF" or cs4 = "FEFEFE"  or cs4 = "FEFFFE" or cs4 = "FEFFFF"  or cs4 = "FFFFFE" or cs4 = "FFFEFE"  or cs4 = "FEFEFF"  Then 
    Else 
        MoveTo 1455,618
        Delay 100
        LeftClick 1
    End If
    If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
        MoveTo 1533,618
        Delay 100
        LeftClick 1
        MoveTo 1633,618
        Delay 100
        LeftClick 1
        Delay 100
        SayString "无"
    Else 
        //    MoveTo 1533,618
        //    Delay 100
        //    LeftClick 1
    End If
End Sub
Sub 住宅类()
    //跳表四
    MoveTo 1546,221
    Delay 100
    LeftClick 1
    Delay 200
    crydsx = GetPixelColor(1278,378)
    If crydsx = "FFFFFF" or crydsx = "FEFEFE"  or crydsx = "FEFFFE" or crydsx = "FEFFFF" or crydsx = "FFFFFE" Then 
    Else 
        MoveTo 1424,373
        Delay 100
        LeftClick 1
        Delay 100
        MoveTo 1413,439
        Delay 100
        LeftClick 1
    End If
    //用能情况
    cn1 = GetPixelColor(1266,580)
    cn2 = GetPixelColor(1329,580)
    cn3 = GetPixelColor(1378,580)
    cn4 = GetPixelColor(1455,580)
    cs1 = GetPixelColor(1266,621)
    cs2 = GetPixelColor(1329,621)
    cs3 = GetPixelColor(1378,621)
    cs4 = GetPixelColor(1455,621)
    //采暖其他
    cr341 = GetPixelColor(1760,588)
    cr342 = GetPixelColor(1759,566)
    cr343 = GetPixelColor(1732,577)
    //炊事其他
    cr351 = GetPixelColor(1757,627)
    cr352 = GetPixelColor(1757,608)
    cr353 = GetPixelColor(1732,617)
    If cn1 = "FFFFFF" or cn1 = "FEFEFE"  or cn1 = "FEFFFE" or cn1 = "FEFFFF" or cn1 = "FFFFFE" or cn1 = "FFFEFE"  or cn1 = "FEFEFF"   Then 
    Else 
        MoveTo 1266,577
        Delay 100
        LeftClick 1
    End If
    If cn2 = "FFFFFF" or cn2 = "FEFEFE"  or cn2 = "FEFFFE" or cn2 = "FEFFFF" or cn2 = "FFFFFE" or cn2 = "FFFEFE"  or cn2 = "FEFEFF"   Then 
        MoveTo 1329,577
        Delay 100
        LeftClick 1
    Else 
        //        MoveTo 1329,577
        //        Delay 100
        //        LeftClick 1
    End If
    If cn3 = "FFFFFF" or cn3 = "FEFEFE"  or cn3 = "FEFFFE" or cn3 = "FEFFFF" or cn3 = "FFFFFE" or cn3 = "FFFEFE"   or cn3 = "FEFEFF"  Then 
    Else 
        MoveTo 1378,577
        Delay 100
        LeftClick 1
    End If
    If cn4 = "FFFFFF" or cn4 = "FEFEFE"  or cn4 = "FEFFFE" or cn4 = "FEFFFF" or cn4 = "FFFFFE" or cn4 = "FFFEFE"   or cn4 = "FEFEFF"  Then 
    Else 
        MoveTo 1455,577
        Delay 100
        LeftClick 1
    End If
    If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
        //        MoveTo 1533,577
        //        Delay 100
        //        LeftClick 1
        //        MoveTo 1633,577
        //        Delay 100
        //        LeftClick 1
        //        Delay 100
        //        SayString "无"
    Else 
        MoveTo 1533,577
        Delay 100
        LeftClick 1
    End If
    If cs1 = "FFFFFF" or cs1 = "FEFEFE"  or cs1 = "FEFFFE" or cs1 = "FEFFFF"  or cs1 = "FFFFFE" or cs1 = "FFFEFE"  or cs1 = "FEFEFF"  Then 
    Else 
        MoveTo 1266,618
        Delay 100
        LeftClick 1
    End If
    If cs2 = "FFFFFF" or cs2 = "FEFEFE"  or cs2 = "FEFFFE" or cs2 = "FEFFFF" or  cs2 = "FFFFFE" or cs2 = "FFFEFE" or cs2 = "FEFEFF"  Then 
        MoveTo 1329,618
        Delay 100
        LeftClick 1
    Else 
        //        MoveTo 1329,618
        //        Delay 100
        //        LeftClick 1
    End If
    If cs3 = "FFFFFF" or cs3 = "FEFEFE"  or cs3 = "FEFFFE" or cs3 = "FEFFFF"  or cs3 = "FFFFFE" or cs3 = "FFFEFE"  or cs3 = "FEFEFF"  Then 
        MoveTo 1378,618
        Delay 100
        LeftClick 1
    Else 
        //        MoveTo 1378,618
        //        Delay 100
        //        LeftClick 1
    End If
    If cs4 = "FFFFFF" or cs4 = "FEFEFE"  or cs4 = "FEFFFE" or cs4 = "FEFFFF"  or cs4 = "FFFFFE" or cs4 = "FFFEFE"  or cs4 = "FEFEFF"  Then 
    Else 
        MoveTo 1455,618
        Delay 100
        LeftClick 1
    End If
    If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
        //        MoveTo 1533,618
        //        Delay 100
        //        LeftClick 1
        //        MoveTo 1633,618
        //        Delay 100
        //        LeftClick 1
        //        Delay 100
        //        SayString "无"
    Else 
        MoveTo 1533,618
        Delay 100
        LeftClick 1
    End If
End Sub
Sub 通过()
    MoveTo 1148, 221
    Delay 100
    LeftClick 1
    Delay 100
    MoveTo 1847,615
    Delay 100
    LeftClick 1
    Delay 100
    MouseWheel - 50 
    //    点击 
    Delay 500
    IfColor 1266, 763, "AA9C01", 2 Then
    Else 
        MoveTo 1266,763
        Delay 100
        LeftClick 1
    End If
    //提交
    MoveTo 1472,945
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
    Delay 700
    IfColor 1070, 485, "F7F7F7", 0 Then
        MoveTo 999,618
        Delay 100
        LeftClick 1
    Else
    End If
    cr1 = GetPixelColor(1082,633)
    Do Until cr1 ="968A00"  
        Delay 200
        cr1 = GetPixelColor(1082,633)  
    Loop
    //调查
    cr000 = GetPixelColor(1888,357)
    Do Until cr000 ="F9F9F9"  
        Delay 200
        cr000 = GetPixelColor(1888,357)  
    Loop
    MoveTo 1702,400
    cr00 = GetPixelColor(1896,413)
    Do Until cr00 ="ECECEC"  
        Delay 200
        cr00 = GetPixelColor(1896,413)  
    Loop
    Delay 200
    LeftClick 1
    RestartScript
End Sub
Sub 翻页()
    x = 1
    Do Until 控制 <> 96
        If 控制 = 96 Then 
            余 = x Mod 4
            Select Case 余 
            Case 1
                //表二
                MoveTo 1255, 221
                Delay 100
                LeftClick 1
            Case 2
                //表三
                MoveTo 1375,221
                Delay 100
                LeftClick 1
            Case 3
                //表四
                MoveTo 1539,221
                Delay 100
                LeftClick 1
                Delay 300
            Case 0
                // 表一  
                MoveTo 1148, 221
                Delay 100
                LeftClick 1
            Case Else
            End Select
            x = x + 1
        Else 
        End If
        控制 = WaitKey()
    Loop
End Sub
Sub 基本安全()
    //跳表四
    MoveTo 1539,221
    Delay 100
    LeftClick 1
    Delay 500
    MoveTo 1300,856
    Delay 100
    LeftClick 1
    Delay 100
    MoveTo 1311,895
    Delay 100
    LeftClick 1
End Sub
Sub 存在风险()
    //跳表四
    MoveTo 1539,221
    Delay 100
    LeftClick 1
    Delay 500
    MoveTo 1300,856
    Delay 100
    LeftClick 1
    Delay 100
    MoveTo 1311,916
    Delay 100
    LeftClick 1
    Delay 100
    MoveTo 1520, 790
    Delay 50
    LeftClick 1
    Delay 100
    MouseWheel -20
    Delay 100
    MoveTo 1520,803
    Delay 100
    LeftClick 1
End Sub
Sub 查看照片()
    MoveTo 1375,221
    Delay 100
    LeftClick 1
    Delay 500
    MoveTo 1355,725 
    Delay 100
    LeftClick 1
    Delay 2000
    KeyPress "Right", 1
    Delay 1500
    KeyPress "Right", 1
    Delay 1000
    KeyPress "Esc", 1
End Sub
Sub 辅助用房()
    //复制主房编号
    MoveTo 1382, 280
    Delay 50
    LeftDoubleClick 1
    Delay 100
    Call 复制()
    zfbh = Plugin.Sys.GetCLB 
    Delay 100
    //检查辅助用房
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
        MoveTo 1755,732
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,279
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,620
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,266
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        //跳转表格
        MoveTo 1755,608
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,352
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,697
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        //跳转表格
        MoveTo 1755,340
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
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
        MoveTo 1755,683
        Delay 50
        LeftClick 1
        Delay 50
        Call 辅助用房信息获取()
    End If
    Rem 完成
    MoveTo 1868, 816
    Delay 100
    LeftClick 1
    Delay 100
    MouseWheel 20
End Sub
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
        yt = "其他"
    End If
    //表三
    MoveTo 833,264
    Delay 50
    LeftClick 1
    Delay 500
    //审批手续
    crydsx = GetPixelColor(713,422)
    //    crjssx = GetPixelColor(1052,422)
    //    crsgsx = GetPixelColor(713,471)
    //    cryssx = GetPixelColor(1052,471)
    //    crspsx = GetPixelColor(713,520)
    //    crdjsx = GetPixelColor(1052,520)
    If crydsx = "FFFFFF" or crydsx = "FEFEFE"  or crydsx = "FEFFFE" or crydsx = "FEFFFF" or crydsx = "FFFFFE" Then 
    Else 
        MoveTo 863,419 
        Delay 100
        LeftClick 1
        Delay 100
        MoveTo 863,482
        Delay 100
        LeftClick 1
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
    If yt = "厨房" Then 
        If cn1 = "FFFFFF" or cn1 = "FEFEFE"  or cn1 = "FEFFFE" or cn1 = "FEFFFF" or cn1 = "FFFFFE" or cn1 = "FFFEFE"  or cn1 = "FEFEFF"   Then 
        Else 
            MoveTo 701,620
            Delay 100
            LeftClick 1
        End If
        If cn2 = "FFFFFF" or cn2 = "FEFEFE"  or cn2 = "FEFFFE" or cn2 = "FEFFFF" or cn2 = "FFFFFE" or cn2 = "FFFEFE"  or cn2 = "FEFEFF"   Then 
            MoveTo 764,620
            Delay 100
            LeftClick 1
        Else 
        End If
        If cn3 = "FFFFFF" or cn3 = "FEFEFE"  or cn3 = "FEFFFE" or cn3 = "FEFFFF" or cn3 = "FFFFFE" or cn3 = "FFFEFE"   or cn3 = "FEFEFF"  Then 
        Else 
            MoveTo 813,620
            Delay 100
            LeftClick 1
        End If
        If cn4 = "FFFFFF" or cn4 = "FEFEFE"  or cn4 = "FEFFFE" or cn4 = "FEFFFF" or cn4 = "FFFFFE" or cn4 = "FFFEFE"   or cn4 = "FEFEFF"  Then 
        Else 
            MoveTo 890,620
            Delay 100
            LeftClick 1
        End If
        If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
        Else 
            MoveTo 969,620
            Delay 100
            LeftClick 1
        End If
        If cs1 = "FFFFFF" or cs1 = "FEFEFE"  or cs1 = "FEFFFE" or cs1 = "FEFFFF"  or cs1 = "FFFFFE" or cs1 = "FFFEFE"  or cs1 = "FEFEFF"  Then 
        Else 
            MoveTo 701,661
            Delay 100
            LeftClick 1
        End If
        If cs2 = "FFFFFF" or cs2 = "FEFEFE"  or cs2 = "FEFFFE" or cs2 = "FEFFFF" or  cs2 = "FFFFFE" or cs2 = "FFFEFE" or cs2 = "FEFEFF"  Then 
            MoveTo 764,661
            Delay 100
            LeftClick 1
        Else 
        End If
        If cs3 = "FFFFFF" or cs3 = "FEFEFE"  or cs3 = "FEFFFE" or cs3 = "FEFFFF"  or cs3 = "FFFFFE" or cs3 = "FFFEFE"  or cs3 = "FEFEFF"  Then 
            MoveTo 813,661
            Delay 100
            LeftClick 1
        Else 
        End If
        If cs4 = "FFFFFF" or cs4 = "FEFEFE"  or cs4 = "FEFFFE" or cs4 = "FEFFFF"  or cs4 = "FFFFFE" or cs4 = "FFFEFE"  or cs4 = "FEFEFF"  Then 
        Else 
            MoveTo 890,661
            Delay 100
            LeftClick 1
        End If
        If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
        Else 
            MoveTo 1070,659
            Delay 100
            LeftClick 1
            Delay 50
            Call 全选()
            Delay 50
            Call 复制()
            Delay 50
            str = Plugin.Sys.GetCLB
            If str = "无" Then 
                MoveTo 969,661
                Delay 100
                LeftClick 1
            End If
        End If
    Else 
        If cn1 = "FFFFFF" or cn1 = "FEFEFE"  or cn1 = "FEFFFE" or cn1 = "FEFFFF" or cn1 = "FFFFFE" or cn1 = "FFFEFE"  or cn1 = "FEFEFF"   Then 
        Else 
            MoveTo 701,620
            Delay 100
            LeftClick 1
        End If
        If cn2 = "FFFFFF" or cn2 = "FEFEFE"  or cn2 = "FEFFFE" or cn2 = "FEFFFF" or cn2 = "FFFFFE" or cn2 = "FFFEFE"  or cn2 = "FEFEFF"   Then 
        Else 
            MoveTo 764,620
            Delay 100
            LeftClick 1
        End If
        If cn3 = "FFFFFF" or cn3 = "FEFEFE"  or cn3 = "FEFFFE" or cn3 = "FEFFFF" or cn3 = "FFFFFE" or cn3 = "FFFEFE"   or cn3 = "FEFEFF"  Then 
        Else 
            MoveTo 813,620
            Delay 100
            LeftClick 1
        End If
        If cn4 = "FFFFFF" or cn4 = "FEFEFE"  or cn4 = "FEFFFE" or cn4 = "FEFFFF" or cn4 = "FFFFFE" or cn4 = "FFFEFE"   or cn4 = "FEFEFF"  Then 
        Else 
            MoveTo 890,620
            Delay 100
            LeftClick 1
        End If
        If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
            MoveTo 969,620
            Delay 100
            LeftClick 1
            MoveTo 1070,620
            Delay 100
            LeftClick 1
            Delay 100
            SayString "无"
        Else 
        End If
        If cs1 = "FFFFFF" or cs1 = "FEFEFE"  or cs1 = "FEFFFE" or cs1 = "FEFFFF"  or cs1 = "FFFFFE" or cs1 = "FFFEFE"  or cs1 = "FEFEFF"  Then 
        Else 
            MoveTo 701,661
            Delay 100
            LeftClick 1
        End If
        If cs2 = "FFFFFF" or cs2 = "FEFEFE"  or cs2 = "FEFFFE" or cs2 = "FEFFFF" or  cs2 = "FFFFFE" or cs2 = "FFFEFE" or cs2 = "FEFEFF"  Then 
        Else 
            MoveTo 764,661
            Delay 100
            LeftClick 1
        End If
        If cs3 = "FFFFFF" or cs3 = "FEFEFE"  or cs3 = "FEFFFE" or cs3 = "FEFFFF"  or cs3 = "FFFFFE" or cs3 = "FFFEFE"  or cs3 = "FEFEFF"  Then 
        Else 
            MoveTo 813,661
            Delay 100
            LeftClick 1
        End If
        If cs4 = "FFFFFF" or cs4 = "FEFEFE"  or cs4 = "FEFFFE" or cs4 = "FEFFFF"  or cs4 = "FFFFFE" or cs4 = "FFFEFE"  or cs4 = "FEFEFF"  Then 
        Else 
            MoveTo 890,661
            Delay 100
            LeftClick 1
        End If
        If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
            MoveTo 969,661
            Delay 100
            LeftClick 1
            MoveTo 1070,659
            Delay 100
            LeftClick 1
            Delay 100
            SayString "无"
        Else        
        End If
    End If
    //取消
    MoveTo 919,875
    Delay 50
    LeftClick 1
    Delay 500
    MoveTo 995,872
    Delay 50
    LeftClick 1
    Delay 500
End Sub