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
MacroID=25e8bd34-dd19-4dc8-9416-8e7a01e58155
Description=°´¼üÐÞ¸Ä2
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
aj = WaitKey()
Do While True
    If aj = 81 Then 
        GetCursorPos x,y
        MoveTo 1852, 222
        Delay 200
        LeftClick 1
        MoveTo x, y
    ElseIf aj = 87 Then
        GetCursorPos x,y
        MoveTo 1900, 600
        Delay 200
        MouseWheel 10 
        Delay 200
        MoveTo 1411, 272
        Delay 200
        LeftClick 1
        MoveTo x, y
    ElseIf aj = 97 Then
        MoveTo 1900, 600
        Delay 200
        MouseWheel 10 
        Delay 200
        MoveTo 1411, 272
        Delay 200
        LeftClick 1
    End If
    aj = WaitKey()
Loop
