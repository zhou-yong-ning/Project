
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
cr1=GetPixelColor(1898,118)
Do Until cr1 ="77472E"  
    Delay 200
    cr1 = GetPixelColor(1898,118)  
Loop
Delay time1*3
MoveTo 214,249
Delay time1
LeftClick 1
//弹出界面操作
cr1=GetPixelColor(1896,250)
Do Until cr1 ="3AC267"  
    Delay 200
    cr1 = GetPixelColor(1896,250)  
Loop
Delay time1 * 2
//点击搜索框
MoveTo 457,248
Delay time1
LeftClick 1
Delay time1
Call 全选()
Delay time1
Call 粘贴()
Delay time1
//点击查询
MoveTo 1443,251
Delay time1
LeftClick 1
//查询结果点击
cr1=GetPixelColor(1856,361)
Do Until cr1 ="3AC267"  
    Delay 200
    cr1 = GetPixelColor(1856,361)  
Loop
Delay time1 * 5
//点击详情
MoveTo 1893,364
Delay time1*2
LeftClick 1
//等待弹出窗口
cr1=GetPixelColor(1874,241)
Do Until cr1 ="FF9E40"  
    Delay 200
    cr1 = GetPixelColor(1874,241)  
Loop
Delay time1*2
//查看照片
MoveTo 1201,308
Delay time1
LeftClick 1
//点击照片
Delay time1 *2
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
        //上键  
    ElseIf aj = 38 Then
        //双击影像
        MoveTo 548,637
        Delay time1
        LeftDoubleClick 2
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
KeyPress "Right", 1
Delay time1
SayString text
Delay time1
KeyPress "Down", 1
Delay time1
KeyPress "Enter", 1
Delay time1
KeyPress "Left", 1
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
MacroID=13e813b3-e628-4f5d-8ec1-9d168febadcf
