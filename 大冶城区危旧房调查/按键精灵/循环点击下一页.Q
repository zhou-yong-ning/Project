[General]
SyntaxVersion=2
BeginHotkey=40
BeginHotkeyMod=8
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=37
StopHotkeyMod=8
RunOnce=1
EnableWindow=
MacroID=7a0597ae-ccfe-4a72-ab9b-da0d66338e1d
Description=循环点击下一页
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

[Attachment]
UEsDBBQAAgAIAI4JA1dntbRRiAAAAJYEAAAKABEAz8LSu9KzLmJtcFVUDQAHB//KZBL/ymQH/8pkc/KdxsIABmZArAHEQkAsAsSMDBJg8QSoPDL4+uULGWhUH5H63nz5Urv43KK6zA9PTpJo3/EH79MqZy1rSP2KWysOdx6++z6tbMrqptSvz86R6L99QK2lU1a3ZH59eYnEcFl26ll2esW9bW1fv7wn0b4NpNl3+AEZ/gOGZzbJ4UlB/I3mB5A+AFBLAQIXCxQAAgAIAI4JA1dntbRRiAAAAJYEAAAKAAkAAAAAAAAAIACAgQAAAADPwtK70rMuYm1wVVQFAAcH/8pkUEsFBgAAAAABAAEAQQAAAMEAAAAAAA==


[Script]
Delay 500
//查找颜色
FindColor 1574, 955, 1893, 979, "968A00", intX, intY
If intX > 1000 And intY > 900 Then 
    MoveTo intX + 50, intY + 10
    Delay 200
    LeftClick 1
End If



