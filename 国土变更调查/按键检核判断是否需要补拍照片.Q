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
MacroID=a6a1d6d9-a6c5-4583-93e8-ba9f2e9761cc
Description=按键检核判断是否需要补拍照片
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
time1 = 10
//查看照片
cr1=GetPixelColor(1270,355)
Do Until cr1 ="FF9E40"  
    Delay 200
    cr1 = GetPixelColor(1270,355)  
Loop
Delay time1*50
MoveTo 1201,308
Delay time1
LeftClick 1
//点击照片
Delay time1 *10
MoveTo 1127,450
Delay time1
LeftClick 1
//按键控制
aj = WaitKey()
Do While True
    //数字1
    If aj = 97 Then 
        text = "1"
        Goto over
        //数字2
    ElseIf aj = 98 Then
        text = "2"
        Goto over
        //数字3   
    ElseIf aj = 99 Then
        text = "3"
        Goto over
        //左键  
    ElseIf aj = 37 Then
        MoveTo 1094,682
        Delay time1
        LeftClick 1
        //右键  
    ElseIf aj = 39 Then
        MoveTo 1860,682
        Delay time1
        LeftClick 1
        //放大影像
    ElseIf aj = 38 Then       
        MoveTo 548,637
        Delay time1
        LeftDoubleClick 2
		//缩小影像  
    ElseIf aj = 40 Then       
        MoveTo 548,637
        Delay time1
        MouseWheel -1
    End If
    aj = WaitKey()
Loop
Rem over
//任务栏图标四
MoveTo 170, 1057
Delay time1
LeftClick 1
//右方向粘贴，左键下键
Delay time1 * 2
SayString text
Delay time1
KeyPress "Down", 1
Delay time1
KeyPress "Enter", 1
Delay time1
//任务栏图标三
MoveTo 122, 1057
Delay time1
LeftClick 1
Delay time1
//点击下一条
MoveTo 1846,250
Delay time1
LeftClick 1
Delay time1*20
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
[General]
MacroID=a6a1d6d9-a6c5-4583-93e8-ba9f2e9761cc
