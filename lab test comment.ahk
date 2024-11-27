;#Requires AutoHotkey v1.0
;Please open and save with Big5

` & space::
Click, 395, 111
Sleep 100
Click, 72, 751
Click, 72, 751
return


` & d::
hiv_comment := "建議HIV RNA 病毒量應重覆作二次作為基礎，若尚未治療，建議每3-4個月作一次；若開始治療，建議第一個月要作一次，爾後3-4個月做一次，後來可4-6個月作一次。"
Clipboard := hiv_comment
InputBox, sp_number, Specimen Number, Please enter the specimen numbers:,
Loop, %sp_number%
{
    Click, 790, 665 ;select the comment field
    Send ^v ; send HIV comment
    Sleep 100
    Click, 395, 111 ; verify the specimen
    Sleep 100
    Click, 73, 220
    Click, 73, 220 ;double click the button of next specimen
}
return