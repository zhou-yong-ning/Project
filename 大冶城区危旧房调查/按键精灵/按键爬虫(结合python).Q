[General]
SyntaxVersion=2
BeginHotkey=39
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=37
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=e98c6a20-bf10-445d-8a3f-8dea6f22cbbe
Description=按键爬虫(结合python)
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
time1 = 100
//复制表格编号
Delay time1*3
Call 复制()
Delay time1
//任务栏图标三
MoveTo 122, 1057
Delay time1
LeftClick 1
Delay time1
//点击查询
cr1=GetPixelColor(1720,130)
Do Until cr1 ="968A00"  
    Delay 200
    cr1 = GetPixelColor(1720,130)  
Loop
Delay time1*3
MoveTo 1797, 363
Delay time1
LeftClick 1
//弹出界面操作
cr1=GetPixelColor(342,286)
Do Until cr1 ="3CA2E6"  
    Delay 200
    cr1 = GetPixelColor(342,286)  
Loop
Delay time1 * 2
//点击搜索框
MoveTo 582,605
Delay time1
LeftClick 1
Delay time1
Call 全选()
Delay time1
Call 粘贴()
Delay time1
//点击查询
MoveTo 861,898
Delay time1
LeftClick 1
//查询结果点击
cr1=GetPixelColor(1720,130)
Do Until cr1 ="968A00"  
    Delay 200
    cr1 = GetPixelColor(1720,130)  
Loop
Delay time1 * 3
//点击详情
MoveTo 1868,412
Delay time1
LeftClick 1
//等待弹出窗口
cr1=GetPixelColor(523,362)
Do Until cr1 ="968A00"  
    Delay 200
    cr1 = GetPixelColor(523,362)  
Loop
Delay time1*2
//关闭窗口
MoveTo 1412,272
Delay time1
LeftClick 1
//等待关闭
cr1=GetPixelColor(1720,130)
Do Until cr1 ="968A00"  
    Delay 200
    cr1 = GetPixelColor(1720,130)  
Loop
//任务栏图标四
MoveTo 170, 1057
Delay time1
LeftClick 1
Delay time1 * 2
KeyPress "Down", 1

Sub 全选()
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
End Sub
Sub 复制()
    KeyDown 17, 1
    KeyPress 67, 1
    KeyUp 17, 1
End Sub
Sub 粘贴()
    KeyDown 17, 1
    KeyPress 86, 1
    KeyUp 17, 1
End Sub

