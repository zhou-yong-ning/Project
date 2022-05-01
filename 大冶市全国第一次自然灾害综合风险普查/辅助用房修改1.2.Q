[General]
SyntaxVersion=2
BeginHotkey=38
BeginHotkeyMod=8
PauseHotkey=37
PauseHotkeyMod=8
StopHotkey=35
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=6e5f4158-b2b9-496f-9d1b-f9ed17fe461e
Description=辅助用房修改
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
cr0 = GetPixelColor(1151, 735)
Do Until cr0 ="FB803B"  
    Delay 200
    cr0 = GetPixelColor(1151, 735)  
Loop
Delay 300
IfColor 1339, 805, "FFFFFF", 0 Then
    Goto 完成
Else 
    Delay 200
    MoveTo 1755,732
    Delay 200
    LeftClick 1
    Call 辅助房屋检查()
End If
Delay 500
//滚动
MoveTo 1888,456
Delay 100
LeftClick 1
Delay 100
MouseWheel - 8 
Delay 200
//2
IfColor 1151, 280, "FB803A", 0 Then
    MoveTo 1755,276
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
    Delay 500
End If
//3
IfColor 1340,695, "FFFFFF", 1 Then
    MoveTo 1755,620
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
Else 
    Goto 完成
End If
//滚动
MoveTo 1888,456
Delay 100
LeftClick 1
Delay 100
MouseWheel - 7
Delay 200
//4
IfColor 1151,267, "FB803A", 0 Then
    MoveTo 1755,265
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
    Delay 500
End If
//5
IfColor 1360,700, "FFFFFF", 1 Then
    MoveTo 1755,610
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
Else 
    Goto 完成
End If
//滚动
MoveTo 1888,456
Delay 100
LeftClick 1
Delay 100
MouseWheel - 6
Delay 200
//6
IfColor 1151,355, "FB803A", 0 Then
    MoveTo 1755,351
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
    Delay 500
End If
//7
IfColor 1360,770, "FFFFFF", 1 Then
    MoveTo 1755,695
    Delay 100
    LeftClick 1
    Call 辅助房屋检查()
Else 
    Goto 完成
End If
//保存并返回
Rem 完成
Delay 200
MoveTo 1395,945
Delay 100
LeftClick 1
//调查
cr000 = GetPixelColor(1828,360)
Do Until cr000 ="F9F9F9"  
    Delay 200
    cr000 = GetPixelColor(1828,360)  
Loop
MoveTo 1815,401
cr00 = GetPixelColor(1855,401)
Do Until cr00 ="ECECEC"  
    Delay 200
    cr00 = GetPixelColor(1855,401)  
Loop
Delay 200
LeftClick 1
Sub 辅助房屋检查()
    cr1 = GetPixelColor(1212,500)
    Do Until cr1 ="EEEEEE"  
        Delay 200
        cr1 = GetPixelColor(1212,500)
    Loop
    Delay 300
    cr2 = GetPixelColor(702,557)
    //跳表三
    MoveTo 833,264
    Delay 200
    LeftClick 1
    crb3 = GetPixelColor(1367,810)
    Do Until crb3 ="F5F5F5"  
        Delay 200
        crb3 = GetPixelColor(1367,810)
    Loop
    Delay 200
    //采暖电  
    cr311 = GetPixelColor(769,616)
    cr312 = GetPixelColor(766,625)
    cr313 = GetPixelColor(762,621)
    //炊事电  
    cr321 = GetPixelColor(761,663)
    cr322 = GetPixelColor(764,665)
    cr323 = GetPixelColor(767,660)
    //炊事气    
    cr331 = GetPixelColor(810,663)
    cr332 = GetPixelColor(813,664)
    cr333 = GetPixelColor(819,657)
    //采暖其他
    cr341 = GetPixelColor(1175,610)
    cr342 = GetPixelColor(1175,630)
    cr343 = GetPixelColor(1190,615)
    //炊事其他
    cr351 = GetPixelColor(1175,650)
    cr352 = GetPixelColor(1173,671)
    cr353 = GetPixelColor(1190,661)
    //用地手续
    MoveTo 863,419 
    Delay 100
    LeftClick 1
    MoveTo 863,482
    Delay 100
    LeftClick 1
    If cr2 = "AA9C00" Then 
        //采暖  
        If cr311 =  "FEFEFE" or cr312 =  "FEFEFE"  or cr313 =  "FEFEFE" or cr311 =  "FEFFFE" or cr312 =  "FEFFFE"  or cr313 =  "FEFFFE" or cr311 =  "FEFFFF" or cr312 =  "FEFFFF"  or cr313 =  "FEFFFF" Then
            MoveTo 766,620
            Delay 100
            LeftClick 1
        Else 
        End If
        //炊事电 
        If cr321 =  "FEFEFE" or cr322 =  "FEFEFE"  or cr323 =  "FEFEFE" or cr321 =  "FEFFFE" or cr322 =  "FEFFFE"  or cr323 =  "FEFFFE" or cr321 =  "FEFFFF" or cr322 =  "FEFFFF"  or cr323 =  "FEFFFF" Then
            MoveTo 766,661
            Delay 100
            LeftClick 1
        Else 
        End If
        //炊事气
        If cr331 =  "FEFEFE" or cr332 =  "FEFEFE" or cr333 =  "FEFEFE" or cr331 =  "FEFFFE" or cr332 =  "FEFFFE"  or cr333 =  "FEFFFE" or cr331 =  "FEFFFF" or cr332 =  "FEFFFF"  or cr333 =  "FEFFFF" Then
            MoveTo 815,661
            Delay 100
            LeftClick 1
        Else 
        End If
        Call 采暖其他无取消()
        Call 炊事其他无取消()
    Else 
        //采暖电 
        Delay 100
        If cr311 =  "FEFEFE" or cr312 =  "FEFEFE"  or cr313 =  "FEFEFE" or cr311 =  "FEFFFE" or cr312 =  "FEFFFE"  or cr313 =  "FEFFFE" or cr311 =  "FEFFFF" or cr312 =  "FEFFFF"  or cr313 =  "FEFFFF" Then
        Else 
            MoveTo 766,620
            Delay 200
            LeftClick 1
        End If
        Delay 100
        //炊事电 
        If cr321 =  "FEFEFE" or cr322 =  "FEFEFE"  or cr323 =  "FEFEFE" or cr321 =  "FEFFFE" or cr322 =  "FEFFFE"  or cr323 =  "FEFFFE" or cr321 =  "FEFFFF" or cr322 =  "FEFFFF"  or cr323 =  "FEFFFF" Then
        Else 
            MoveTo 766,661
            Delay 200
            LeftClick 1
        End If
        //炊事气
        Delay 100
        If cr331 =  "FEFEFE" or cr332 =  "FEFEFE" or cr333 =  "FEFEFE" or cr331 =  "FEFFFE" or cr332 =  "FEFFFE"  or cr333 =  "FEFFFE" or cr331 =  "FEFFFF" or cr332 =  "FEFFFF"  or cr333 =  "FEFFFF" Then
        Else 
            MoveTo 815,661
            Delay 200
            LeftClick 1
        End If
        Call 采暖其他无()
        Call 炊事其他无()
    End If
    //保存
    MoveTo 919, 871
    Delay 200
    LeftClick 1
    Delay 500
    MoveTo 995, 871
    Delay 200
    LeftClick 1
End Sub
Sub 采暖其他无()
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
End Sub
Sub 炊事其他无()
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
End Sub
Sub 采暖其他无取消()
    If cr341 = "EEEEEE" or cr342 = "EEEEEE" or cr343 = "EEEEEE" Then 
    Else 
        MoveTo 969,620
        Delay 100
        LeftClick 1
    End If
End Sub
Sub 炊事其他无取消()
    If cr351 = "EEEEEE" or cr352 = "EEEEEE" or cr353 = "EEEEEE" Then 
    Else 
        MoveTo 969,661
        Delay 100
        LeftClick 1
    End If
End Sub
