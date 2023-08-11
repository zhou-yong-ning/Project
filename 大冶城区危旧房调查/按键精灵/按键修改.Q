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
MacroID=e1c25e00-d71e-488e-af1d-074e901f21ea
Description=°´¼üÐÞ¸Ä
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
    If aj = 68 Then 
        KeyPress "Right", 1
    ElseIf aj = 65 Then
        KeyPress "Left", 1
    ElseIf aj = 83 Then
        KeyPress "Esc", 1
    End If
    aj = WaitKey()
Loop