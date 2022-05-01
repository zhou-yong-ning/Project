[General]
SyntaxVersion=2
BeginHotkey=38
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=37
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=1b899672-0fd1-48ef-9818-f9f7097fefcc
Description=农村房屋信息获取
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
    Call 非住宅()
    fwlb = "非住宅"
    //现场调查情况
    IfColor 1266,690, "AA9C01", 0 Then
        dc = "不需要调查"
    Else
        dc = "需要调查"
    End If
Else 
    Delay 200
    Call 住宅()
    fwlb = "住宅"
    //现场调查情况
    IfColor 1266, 847, "AA9C01", 0 Then
        dc = "不需要调查"
    Else
        dc = "需要调查"
    End If
End If
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
SayString dc
Delay 30
KeyPress "Left", 6
Delay 30
KeyPress "Down", 1
Delay 30
MoveTo 123,1060
Delay 30
LeftClick 1
Delay 100
//调查
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
//510681197902084120