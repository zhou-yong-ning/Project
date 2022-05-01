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
MacroID=4689d6f4-131d-4f0e-9fb5-737ed4a9f38e
Description=预制板问题1.2
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
cr0 = GetPixelColor(1151,282)
Do Until cr0 ="FB803B"  
    Delay 200
    cr0 = GetPixelColor(1151,282)  
Loop
Delay 500
cr0 = GetPixelColor(1151, 282)
IfColor 1266, 330, "FFFFFF", 0 Then
    I= 非住宅
Else 
    I= 住宅
End If
//跳表二  
MoveTo 1256,221
Delay 100
LeftClick 1
//条件判断
cr表二 = GetPixelColor(1463,528)
Do Until cr表二 ="EEEEEE"  
    Delay 200
    cr表二 = GetPixelColor(1463,528)  
Loop
//预制板判断
Delay 200
cry1 = GetPixelColor(1263,626)
cry2 = GetPixelColor(1266,628)
cry3 = GetPixelColor(1271,620)
Delay 100
crjg1 = GetPixelColor(1266, 463)
crjg2 = GetPixelColor(1357, 463)
If crjg1 = "AA9C01" Then 
    IfColor 1357,725,"AA9C01",0 Then
    Else 
        MoveTo 1357, 725
        Delay 100
        LeftClick 1
    End If
ElseIf crjg2 = "AA9C01" Then
    IfColor 1357,624,"AA9C01",0 Then
    Else 
        MoveTo 1357,624
        Delay 100
        LeftClick 1
    End If
Else 
    IfColor 1357,580,"AA9C01",0 Then
    Else 
        MoveTo 1357,580
        Delay 100
        LeftClick 1
    End If
End If
//跳表四
MoveTo 1519, 224
Delay 50
LeftClick 1
Delay 100
//条件判断
cr表四 = GetPixelColor(1266,279)
Do Until cr表四 = "AA9C01"  
    Delay 200
    cr表四 = GetPixelColor(1266,279)  
Loop
//用地手续
MoveTo 1431,377
Delay 50
LeftClick 1
MoveTo 1431,443
Delay 30
LeftClick 1
//安全隐患判断
If crjg1 = "AA9C01" Then 
    If cry1 = "FFFFFF" or cry2 = "FFFFFF" or cry3 = "FFFFFF" Then 
        Call 基本安全()
    Else 
        MoveTo 1300,856
        Delay 100
        LeftClick 1
        Delay 100
        IfColor 1393,916,"FF901E",1 Then 
            MoveTo 1311,916
            Delay 50 
            LeftClick 1
            Delay 50
            MoveTo 1520, 790
            Delay 50
            LeftClick 1
            Delay 50
            MouseWheel -20
            Delay 100
            MoveTo 1520,803
            Delay 200
            LeftClick 1
        End If
    End If
Else 
    Call 基本安全()
End If
//保存返回
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
MoveTo 1815, 405
cr00 = GetPixelColor(1855, 405)
Do Until cr00 ="ECECEC"  
    Delay 200
    cr00 = GetPixelColor(1855, 405)  
Loop
Delay 200
LeftClick 1
Sub 基本安全()
    MoveTo 1418,852 
    Delay 100
    LeftClick 1
    Delay 200
    IfColor 1414,898,"FF901E",0 Then
    Else 
        MoveTo 1414,898
        Delay 100
        LeftClick 1
    End If
End Sub
