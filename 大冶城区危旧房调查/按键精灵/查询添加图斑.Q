[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=d00fc687-9234-4ccc-9109-4badc900643f
Description=查询添加图斑
Enable=1
AutoRun=0
[Repeat]
Type=0
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
MoveTo 342,605
Delay time1
LeftClick 1
Delay time1
Call 全选()
Delay time1
Call 粘贴()
Delay time1
MoveTo 861,898
Delay time1
LeftClick 1
//查询结果添加
cr1=GetPixelColor(1720,130)
Do Until cr1 ="968A00"  
    Delay 200
    cr1 = GetPixelColor(1720,130)  
Loop
Delay time1 * 3
//添加条件判断，防止误添加
IfColor 1739,495,"FFFFFF",1 Then
    ExitScript
End If
MoveTo 1555,362
Delay time1
LeftClick 1
MoveTo 1868,364
Delay time1
LeftClick 1
Delay time1 * 5
KeyPress "Enter", 1
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

