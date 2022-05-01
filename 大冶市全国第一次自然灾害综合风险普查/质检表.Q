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
MacroID=8da541fe-0845-445a-b3c8-79f9e8b860c9
Description=质检表
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
//安全措施
crchushi = GetPixelColor(1151, 283)
Do Until crchushi ="FB803B"  
    Delay 200
    crchushi = GetPixelColor(1151, 283)  
Loop
Delay 200
IfColor 1266, 330, "FFFFFF", 0 Then
    I= "非住宅"
Else 
    I = "住宅"
End If
If I = "住宅" Then 
    crzhu = GetPixelColor(1151, 735)
    Do Until crzhu ="FB803B"  
        Delay 200
        crzhu = GetPixelColor(1151, 735)  
    Loop
    Delay 300
ElseIf I = "非住宅" Then
    crfzhu = GetPixelColor(1637,589)
    Do Until crfzhu ="AA9C01"  
        Delay 200
        crfzhu = GetPixelColor(1637,589)  
    Loop
    Delay 300
End If
//复制主房编号
MoveTo 1382, 280
Delay 50
LeftDoubleClick 1
Delay 50
Call 复制()
zfbh = Plugin.Sys.GetCLB 
zfbhs = CStr(zfbh)
MoveTo 172,1056
Delay 50
LeftClick 1
Delay 200
SayString zfbhs
Delay 100
KeyPress "Down", 1
Delay 50
MoveTo 125,1056
Delay 50
LeftClick 1
Delay 200
//检查辅助用房
If I = "住宅" Then 
    MoveTo 1382,735
    Delay 50
    LeftDoubleClick 1
    Delay 50
    Call 复制()
    Delay 100
    fzfbh1 = Plugin.Sys.GetCLB
    Delay 50
    fzfbhs1 = CStr(fzfbh1)
    Delay 50
    If fzfbhs1 = zfbhs Then 
        num = 0
        Goto 结束if
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        KeyPress "Right", 1
        Delay 50
        SayString fzfbhs1
        Delay 50
        KeyPress "Down", 1
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200      
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
    fzfbhs2 = CStr(fzfbh2)
    Delay 50
    If fzfbhs2 = fzfbhs1 or Len(fzfbhs2)>16 Then 
        num = 1
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs2
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    fzfbhs3 = CStr(fzfbh3)
    Delay 50
    If fzfbhs3 = fzfbhs2 Then 
        num = 2
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs3
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    If fzfbhs4 = fzfbhs3 Then 
        num = 3
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs4
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    fzfbhs5 = CStr(fzfbh5)
    Delay 50
    If fzfbhs5 = fzfbhs4 Then 
        num = 4
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs5
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    fzfbhs6 = CStr(fzfbh6)
    Delay 50
    If fzfbhs6 = fzfbhs5 Then 
        num = 5
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs6
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    fzfbhs7 = CStr(fzfbh7)
    Delay 50
    If fzfbhs7 = fzfbhs6 Then 
        num = 6
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs7
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
    End If
    //滚动
    MoveTo 1888,456
    Delay 100
    LeftClick 1
    Delay 100
    MouseWheel - 7
    Delay 200
    //8
    MoveTo 1381,340
    Delay 50
    LeftDoubleClick 1
    Delay 50
    Call 复制()
    Delay 50
    fzfbh8 = Plugin.Sys.GetCLB
    Delay 50
    fzfbhs8 = CStr(fzfbh8)
    Delay 50
    If fzfbhs8 = fzfbhs7 Then 
        num = 7
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs8
        Delay 50
        KeyPress "Down", 1
        Delay 50
        MoveTo 125,1056
        Delay 50
        LeftClick 1
        Delay 200 
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
    fzfbhs9 = CStr(fzfbh9)
    Delay 50
    If fzfbhs9 = fzfbhs8 Then 
        num = 8
        Goto 完成
    Else 
        //跳转表格
        MoveTo 172,1056
        Delay 50
        LeftClick 1
        Delay 200
        SayString fzfbhs9
        Delay 50
        KeyPress "Down", 1
        Delay 50
        KeyPress "Left", 1
        Goto 结束if
    End If
    Rem 完成
    MoveTo 172,1056
    Delay 50
    LeftClick 1
    Delay 200
    KeyPress "Left", 1
    Delay 50
    MoveTo 123,1056
    Delay 50
    LeftClick 1
End If
Rem 结束if
//保存返回
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
Sub 复制()
    KeyDown 17, 1
    KeyPress 67, 1
    KeyUp 17, 1
End Sub