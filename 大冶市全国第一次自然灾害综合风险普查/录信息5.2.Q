[General]
SyntaxVersion=2
BeginHotkey=38
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=35
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=66a79b51-c9fc-447e-90fc-2274f8ab9b21
Description=5.1
Enable=1
AutoRun=0
[Repeat]
Type=1
Number=1
[SetupUI]
Type=1
QUI=Form1
[Relative]
SetupOCXFile=
[Comment]

[Attachment]
UEsDBBQAAgAIAPJpLVRJc/oHhQEAAIYFAAAFABEAMS5ibXBVVA0AB+0l4GH8JeBhdYjbYaWRy0oCURjHJ8KICEGirAiJIOwJbFEr2wlSEbgvCop6jl6gtItk9xIKXKiUiy5eWrVLe4RGfYBRz5w55/Tp0WGcG1h/fmcx/+/7fpvxB/cdQis+eHPwJuG54PUJ461+pT3XhjHmjNb6I//FHauDquVbf0GOo5oNa88olEEDtjsbr6jjq0h0LFYfPDZnOd1chIQystUOnFdrtOODHHzhoZO6KYFU27f6hKx2DosK3+n4ZMJ8D2j4tGFKICUH07LVdP4RYcK6fZC8SJzRxh8olIkq0fggm2/YdYaMfFYpYDraesdaQ7evLFHPNRqJ6eFTYw/LcGLtg4SLZPRc1sFHxj5cIrpzgw9+zEICuy9kLXykKxcTWNbrjD5ITqQTl1gL73VlXqTGWzMfZDurTF1hFV5qm52sYnpo4fuRmDeOPTdteKl+wkiUWC++5o8p0elbhfNRoYD6GfmmVlfWPkSYP6nM3OlZSiqIsN59kKxIZ++JjpxIbU5sfZDdPPXGicpegdrvC8IvUEsBAhcLFAACAAgA8mktVElz+geFAQAAhgUAAAUACQAAAAAAAAAgAICBAAAAADEuYm1wVVQFAAftJeBhUEsFBgAAAAABAAEAPAAAALkBAAAAAA==


[Script]
//镇名与村名
zhen = "金湖街道"
cun = "柯庄村村民委员会"
//(调查中/未调查)
调查 = "调查中"  
//调查 = "未调查"
//安全设置
cr1 = GetPixelColor(1151, 652)
Do Until cr1 ="FB803B"  
    Delay 200
    cr1 = GetPixelColor(1151, 652)  
Loop
Delay 200
//开始执行
MoveTo 1351, 281
Delay 100
LeftDoubleClick 1
'双击复制
Delay 100
Call 复制()
Delay 100
fwbh = Plugin.Sys.GetCLB 
//表格一基本信息
MoveTo 1265, 373
Delay 150
LeftClick 1
//镇
//MoveTo 1314, 460
//Delay 100
//LeftClick 1
//cr2 = GetPixelColor(1377, 531)
//i = 1
//Do Until cr2 ="B77A33"  
//    Delay 100
//    cr2 = GetPixelColor(1377, 531)
//    i = i + 1
//    If i=20 Then  
//        Call 刷新网页()
//    End If 
//Loop
//Delay 150
//SayString zhen
//Delay 50
//KeyPress 13,1
//村
//MoveTo 1551, 456
//Delay 100
//LeftClick 1
//cr3 = GetPixelColor(1581,531)
//Do Until cr3 ="B77A33"  
//    Delay 100
//    cr3 = GetPixelColor(1581,531)  
//Loop
//Delay 100
//SayString cun
//Delay 100
//KeyPress 13, 1
//Delay 50
//MoveTo 1643,584
//Delay 100
//LeftClick 1
//选择表格 查询
MoveTo 172, 1059
Delay 100
LeftClick 1
Delay 200
//调出查找框
KeyDown 17, 1
KeyPress 70, 1
KeyUp 17, 1
Delay 100
KeyPress "BackSpace", 20
//粘贴
Delay 100
Call 粘贴()
Delay 100
//回车
KeyDown "Enter", 1
Delay 100
KeyUp "Enter", 1
Delay 300
//复制宗地码
KeyPress 27, 1
Delay 100
KeyPress 39, 1
Delay 50
Call 复制()
Delay 50
房屋类型 = Plugin.Sys.GetCLB()
Delay 50
fwlx = Mid(房屋类型, 13, 2)
If fwbh = 房屋类型 Then 
    KeyPress 27, 1
    Delay 100
    Call 无宗地码()
    结构 = 5
    Goto 表三、四
End If
//获取姓名
KeyPress 39, 1
Delay 50
Call 复制()
Delay 50
姓名 = Plugin.Sys.GetCLB()
//获取身份证
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
身份证 = Plugin.Sys.GetCLB()
//获取小区
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
小区 = Plugin.Sys.GetCLB()
//获取组
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
组 = Plugin.Sys.GetCLB()
//获取号
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
号 = Plugin.Sys.GetCLB()
//获取层
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
楼层 = Plugin.Sys.GetCLB()
//获取面积
KeyPress 39,1
Delay 50
Call 复制()
Delay 50
面积 = Plugin.Sys.GetCLB()
//判断是否住宅
If fwlx = "JB" Then 
    Call 公共设施类()
    Goto 表三、四
End If
//填写姓名
MoveTo 123, 1059
Delay 30
LeftClick 1
Delay 100
MoveTo 1311, 582
Delay 100
LeftClick 1
Delay 100
Call 全选()
Delay 50
SayString 姓名
Delay 30
//填写身份证
MoveTo 1311, 632
Delay 50
LeftClick 1
Delay 30
Call 全选()
Delay 50
SayString 身份证
Delay 30
//填写小区
MoveTo 1480, 511
Delay 50
LeftClick 1
Delay 30
Call 全选()
Delay 50
SayString 小区
Delay 30
//填写组
MoveTo 1310, 512
Delay 50
LeftClick 1
Delay 30
Call 全选()
Delay 50
SayString 组
Delay 30
//填写号
MoveTo 1720, 512
Delay 50
LeftClick 1
Delay 50
Call 全选()
Delay 50
SayString 号
Delay 100
//填写楼层
MoveTo 1256, 220
Delay 50
LeftClick 1
Delay 300
MoveTo 1280,280
Delay 50
LeftClick 1
Delay 50
Call 全选()
Delay 50
SayString 楼层
//填写面积
MoveTo 1280, 330
Delay 50
LeftClick 1
Delay 50
Call 全选()
Delay 50
SayString 面积
//建造方式，安全鉴定
//MoveTo 1357,580
//Delay 200
//LeftClick 1
//MoveTo 1415,670
//Delay 200
//LeftClick 1
//MoveTo 1410,730
//Delay 200
//LeftClick 1
//填年份
//If 年份 <= 1980 Then 
//    MoveTo 1265,390
//    Delay 100
//    LeftClick 1
//    Delay 100
//ElseIf 年份>=1981 and 年份<=1990 Then  
//    MoveTo 1388,390
//    Delay 100
//    LeftClick 1
//    Delay 100
//ElseIf 年份>=1991 and 年份<=2000 Then
//    MoveTo 1504,390
//    Delay 100
//    LeftClick 1
//    Delay 200
//ElseIf 年份>=2001 and 年份<= 2010 Then
//    MoveTo 1266,419
//    Delay 100
//    LeftClick 1
//    Delay 100
//ElseIf 年份 >= 2011 and 年份 <= 2015 Then  
//    MoveTo 1382,419
//    Delay 100
//    LeftClick 1
//    Delay 100
//ElseIf 年份 >= 2016 and 年份 <=2022 Then      
//    MoveTo 1498,419
//    Delay 100
//    LeftClick 1
//    Delay 100
//Else 
//End If
//判断结构类型
//Select Case 结构
//Case 5
//    //砖石结构预制版
//    MoveTo 1266,463
//    Delay 150
//    LeftClick 1
//    MoveTo 1266,580
//    Delay 150
//    LeftClick 1
//    MoveTo 1268,623
//    Delay 150
//    LeftClick 1
//    MoveTo 1319,665 
//    Delay 150
//    LeftClick 1
//Case 6
//    //生土结构
//    MoveTo 1357,463
//    Delay 150
//    LeftClick 1
//    MoveTo 1266,580
//    Delay 150
//    LeftClick 1
//Case 9
//    //砖石结构木或轻钢楼屋盖
//    MoveTo 1266,463
//    Delay 150
//    LeftClick 1
//    MoveTo 1266,580
//    Delay 150
//    LeftClick 1
//    MoveTo 1422,623
//    Delay 150
//    LeftClick 1
//    MoveTo 1319,665 
//    Delay 150
//    LeftClick 1
//Case Else
//End Select
//Rem 表三、四
//抗震设防信息
//MoveTo 1369,224
//Delay 100
//LeftClick 1
//条件判断
//Delay 200
//MoveTo 1343, 281
//Delay 80
//LeftClick 1
//MoveTo 1345, 325
//Delay 80
//LeftClick 1
//MoveTo 1343, 386
//Delay 80
//LeftClick 1
//MoveTo 1340, 439
//Delay 80
//LeftClick 1
//添加照片
//IfColor 1213,820,"AA9C00",0 Then
//    MoveTo 1213,820
//    Delay 100
//    LeftClick 1
//    cr照片2 = GetPixelColor(480,90)
//    Do Until cr照片2 ="F7F6F5"  
//        Delay 100
//        cr照片2 = GetPixelColor(480,90)  
//    Loop
//    MoveTo 229,167
//    Delay 100
//    LeftDoubleClick 1
//    Delay 300
//Else    
//End If
//Delay 200
//Rem 表四
//安全隐患排查
//MoveTo 1519, 224
//Delay 50
//LeftClick 1
//Delay 100
//条件判断
//cr3 = GetPixelColor(1266,281)
//Do Until cr3 ="AA9C01"  
//    Delay 200
//    cr3 = GetPixelColor(1266,281)  
//Loop 
//用地手续
//MoveTo 1431,377
//Delay 50
//LeftClick 1
//MoveTo 1431,443
//Delay 30
//LeftClick 1
//施工许可手续
//MoveTo 1431,425
//Delay 10
//LeftClick 1
//MoveTo 1431,495
//Delay 10
//LeftClick 1
//经营审批手续
//MoveTo 1431,476
//Delay 10
//LeftClick 1
//MoveTo 1431,544
//Delay 10
//LeftClick 1
//规划建设手续
//MoveTo 1750,376
//Delay 10
//LeftClick 1
//MoveTo 1750,449
//Delay 10
//LeftClick 1
//竣工验收手续
//MoveTo 1750,426
//Delay 10
//LeftClick 1
//MoveTo 1750,494
//Delay 10
//LeftClick 1
//房屋登记手续
//MoveTo 1750,476
//Delay 20
//LeftClick 1
//MoveTo 1750,545
//Delay 20
//LeftClick 1
//用能情况
//MoveTo 1329,576
//Delay 20
//LeftClick 1
//MoveTo 1331,616
//Delay 20
//LeftClick 1
//MoveTo 1379,617
//Delay 20
//LeftClick 1
//地质灾害隐患
//MoveTo 1303,807
//Delay 20
//LeftClick 1
//MoveTo 1309,876
//Delay 20
//LeftClick 1
//安全隐患判断
//Select Case 结构
//Case 5
//    Call 存在风险()
//Case Else
//    Call 基本安全()
//End Select
//完成
//Call 完成()
//Delay 500
//IfColor 1151,734,"FB803B",0 Then
//    //颜色一样
//    MoveTo 1416,630
//    Delay 30 
//    LeftClick 1
//    Delay 50
//    LeftDoubleClick 1
//    Delay 50
//    SayString "510681197902084120"
//    Delay 100
//    Call 完成
//End If
Rem 表三、四
Call 完成()
//点调查（判断调查条件）
cr1 = GetPixelColor(1124,401 )
Do Until cr1 ="676767"  
    Delay 200
    cr1 = GetPixelColor(1124,401 )  
Loop
//调查
cr000 = GetPixelColor(1870,395)
Do Until cr000 ="F9F9F9"  
    Delay 200
    cr000 = GetPixelColor(1870,395)  
Loop
MoveTo 1823,435
cr00 = GetPixelColor(1895,444)
Do Until cr00 ="ECECEC"  
    Delay 200
    cr00 = GetPixelColor(1895,444)  
Loop
Delay 200
LeftClick 1
Sub 无宗地码()
    MoveTo 123,1059
    LeftClick 1
    Delay 100
    //填路  
    MoveTo 1465, 512
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    SayString "0"
    //填姓名  
    MoveTo 1292, 582
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    SayString "不详"
    //输入固定身份证号
    MoveTo 1300, 630
    Delay 200
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    SayString "510681197902084120"
    Delay 10
    //填组
    MoveTo 1280, 512
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    SayString "1"
    //表二固定值
    //填写楼层
    MoveTo 1256, 220
    Delay 50
    LeftClick 1
    Delay 100
    MoveTo 1280,280
    Delay 100
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 50
    SayString "1"
    Delay 100
    //填写面积
    MoveTo 1280,330
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 50
    SayString "50"
    Delay 100
    //    //建造方式，安全鉴定
    //    MoveTo 1357,580
    //    Delay 200
    //    LeftClick 1
    //    MoveTo 1415,670
    //    Delay 200
    //    LeftClick 1
    //    MoveTo 1410,730
    //    Delay 200
    //    LeftClick 1
    //    //年份默认1980年
    //    MoveTo 1265,390
    //    Delay 100
    //    LeftClick 1
    //    Delay 100
    //    //砖石结构预制版    
    //    MoveTo 1266,463
    //    Delay 200
    //    LeftClick 1
    //    MoveTo 1266,580
    //    Delay 200
    //    LeftClick 1
    //    MoveTo 1268,623
    //    Delay 200
    //    LeftClick 1
    //    MoveTo 1319,665 
    //    Delay 200
    //    LeftClick 1
End Sub
Sub 公共设施类()
    MoveTo 123,1059
    Delay 100
    LeftClick 1
    //改为非住宅
    MoveTo 1334, 331
    Delay 200
    LeftClick 1
    //    cr1 = GetPixelColor(1374,411)
    //    Do Until cr1 ="343434"  
    //        Delay 100
    //        cr1 = GetPixelColor(1374,411)  
    //    Loop
    //    Delay 500
    //    MoveTo 1325, 415
    //    Delay 200
    //    LeftClick 1
    //    Delay 100
    //    cr1 = GetPixelColor(1318,489)
    //    Do Until cr1 ="B77A33"  
    //        Delay 100
    //        cr1 = GetPixelColor(1318,489)  
    //    Loop
    //    Delay 300
    //    //镇    
    //    SayString zhen
    //    Delay 100
    //    KeyPress 13,1
    //    //村  
    //    MoveTo 1560,407 
    //    Delay 100
    //    LeftClick 1
    //    cr1 = GetPixelColor(1561,489)
    //    Do Until cr1 ="B77A33"  
    //        Delay 100
    //        cr1 = GetPixelColor(1561,489)  
    //    Loop
    //    Delay 100
    //    SayString cun
    //    Delay 100
    //    KeyPress 13,1
    //粘贴房屋名    
    MoveTo 1301, 539
    Delay 100
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 100
    SayString 姓名
    Delay 100
    MoveTo 1327, 584
    Delay 100
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 100
    SayString 姓名
    Delay 100
    //    //勾选使用人  
    //    MoveTo 1637, 586
    //    Delay 100
    //    LeftClick 1
    //    Delay 100 
    //填写小区  
    MoveTo 1478,464
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 100
    SayString 小区
    Delay 50
    //填写组
    MoveTo 1316,469 
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 100
    SayString 组
    Delay 100
    //填写号  
    MoveTo 1710,467 
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 100
    SayString 号
    Delay 100
    //填写楼层
    MoveTo 1256, 220
    Delay 100
    LeftClick 1
    Delay 100
    MoveTo 1280,280
    Delay 100
    LeftClick 1
    Delay 100
    Call 全选()
    Delay 100
    SayString 楼层
    Delay 100
    //填写面积
    MoveTo 1280, 330
    Delay 50
    LeftClick 1
    Delay 50
    Call 全选()
    Delay 100
    SayString 面积
    Delay 100
    //    //填写年份
    //    //填年份
    //    If 年份 <= 1980 Then 
    //        MoveTo 1265,390
    //        Delay 100
    //        LeftClick 1
    //        Delay 100
    //    ElseIf 年份>=1981 and 年份<=1990 Then  
    //        MoveTo 1388,390
    //        Delay 100
    //        LeftClick 1
    //        Delay 100
    //    ElseIf 年份>=1991 and 年份<=2000 Then
    //        MoveTo 1504,390
    //        Delay 100
    //        LeftClick 1
    //        Delay 200
    //    ElseIf 年份>=2001 and 年份<= 2010 Then
    //        MoveTo 1266,419
    //        Delay 100
    //        LeftClick 1
    //        Delay 100
    //    ElseIf 年份 >= 2011 and 年份 <= 2015 Then  
    //        MoveTo 1382,419
    //        Delay 100
    //        LeftClick 1
    //        Delay 100
    //    ElseIf 年份 >= 2016 and 年份 <=2022 Then      
    //        MoveTo 1498,419
    //        Delay 100
    //        LeftClick 1
    //        Delay 100
    //    Else 
    //    End If
    //    //是否安全鉴定 
    //    MoveTo 1415,804
    //    Delay 150
    //    LeftClick 1
    //    MoveTo 1410,864
    //    Delay 150
    //    LeftClick 1
    //    //建筑用途  
    //    MoveTo 1450,667
    //    Delay 150
    //    LeftClick 1
    //    //建造方式  
    //    MoveTo 1357,580
    //    Delay 150
    //    LeftClick 1
    //    //结构类型
    //    Select Case 结构
    //    Case 5
    //        //砖石结构预制版(红)
    //        MoveTo 1266,463
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1266,580
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1268,623
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1319,665 
    //        Delay 200
    //        LeftClick 1
    //    Case 6
    //        //生土结构
    //        MoveTo 1357,463
    //        Delay 150
    //        LeftClick 1
    //        MoveTo 1266,580
    //        Delay 150
    //        LeftClick 1
    //    Case 9
    //        //砖石结构木或轻钢楼层（深绿）
    //        MoveTo 1266,463
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1266,580
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1422,623
    //        Delay 200
    //        LeftClick 1
    //        MoveTo 1319,665 
    //        Delay 200
    //        LeftClick 1
    //    Case Else
    //        //空白
    //    End Select
    //    //切换至表三  
    //    MoveTo 1369,224
    //    Delay 150
    //    LeftClick 1
    //    //条件判断
    //    Delay 100
    //    cr1 = GetPixelColor(1212,835)
    //    Do Until cr1 ="AA9C00"  
    //        Delay 200
    //        cr1 = GetPixelColor(1212,835)  
    //    Loop
    //    Delay 200
    //    MoveTo 1343, 281
    //    LeftClick 1
    //    Delay 100
    //    MoveTo 1345, 325
    //    LeftClick 1
    //    Delay 100
    //    MoveTo 1343, 386
    //    LeftClick 1
    //    Delay 100
    //    MoveTo 1340, 439
    //    LeftClick 1
    //    Delay 100
    //    //添加照片
    //    cr照片1 = GetPixelColor(1213,834)
    //    Do Until cr照片1 ="AA9C00"  
    //        Delay 100
    //        cr照片1 = GetPixelColor(1213,834)  
    //    Loop
    //    MoveTo 1213,834
    //    Delay 100
    //    LeftClick 1
    //    cr照片2 = GetPixelColor(480,90)
    //    Do Until cr照片2 ="F7F6F5"  
    //        Delay 100
    //        cr照片2 = GetPixelColor(480,90)  
    //    Loop
    //    MoveTo 229,167
    //    Delay 100
    //    LeftDoubleClick 1
    //    Delay 500
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
Sub 基本安全()
    MoveTo 1300,856
    Delay 20
    LeftClick 1
    MoveTo 1311,895
    Delay 20
    LeftClick 1
End Sub
Sub 存在风险()
    MoveTo 1300,856
    Delay 30
    LeftClick 1
    MoveTo 1311,916
    Delay 30 
    LeftClick 1
    Delay 30
    MoveTo 1520,790
    LeftClick 1
    Delay 50
    MouseWheel -20
    Delay 50
    MoveTo 1520,803
    Delay 100
    LeftClick 1
End Sub
Sub 刷新网页()
    Delay 500
    KeyPress "F5", 1
    Delay 5000
    crwy = GetPixelColor(582,100)  
    Do Until crwy = "FFFFFF"
        KeyPress "F5", 1
        Delay 5000
        crwy = GetPixelColor(582,100)  
    Loop
    MoveTo 541, 99
    Delay 500
    LeftClick 1
    Delay 500
    MoveTo 541, 149
    Delay 500
    LeftClick 1
    Delay 500
    crwy2 = GetPixelColor(1200,340)  
    Do Until crwy2 = "F9F9F9"
        Delay 200
        crwy2 = GetPixelColor(1200,340)  
    Loop
    // 选择村
    Delay 1000
    //  查询  
    MoveTo 1154,209
    Delay 500
    LeftClick 1
    crwy4 = GetPixelColor(1438,405)
    Do Until crwy4 ="AA9C01"  
        Delay 200
        crwy4 = GetPixelColor(1438,405)  
    Loop
    MoveTo 1295,305
    Delay 500
    LeftClick 1
    Delay 500
    查询村 = Left(cun, 3)
    Delay 500
    SayString 查询村
    //  点查询  
    MoveTo 1474,405
    Delay 500
    LeftClick 1
    Delay 500
    //  点村  
    MoveTo 1776, 608
    crwy6 = GetPixelColor(1855,598)
    Do Until crwy6 ="ECECEC"  
        Delay 200
        crwy6 = GetPixelColor(1855,598)  
    Loop   
    Delay 500
    LeftClick 1
    Select Case 调查
    Case "调查中"
        // 选择调查中
        MoveTo 1226, 318
        Delay 1000
        LeftClick 1
        Delay 500
    Case Else
    End Select
    //点调查（判断调查条件）
    cr1 = GetPixelColor(1124,401 )
    Do Until cr1 ="676767"  
        Delay 200
        cr1 = GetPixelColor(1124,401 )  
    Loop
    If 调查 = "调查中" Then 
        调查x = 1819 
    ElseIf 调查 = "未调查" Then
        调查x = 1653
    Else 
        EndScript
    End If
    MoveTo 调查x,401
    Delay 100
    LeftClick 1
    RestartScript
End Sub
Sub 完成()
    MoveTo 1564,945
    Delay 200
    LeftClick 1
    cr3 = GetPixelColor(1082,602 )
    Do Until cr3 <> "968A00"  
        Delay 100
        cr3 = GetPixelColor(1082,602 )  
    Loop
    Delay 200
    MoveTo 920,632
    Delay 100
    LeftClick 1
End Sub
Sub 全选()
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
End Sub
