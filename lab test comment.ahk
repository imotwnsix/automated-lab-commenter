;#Requires AutoHotkey v1.0
;Please open and save with Big5

` & space::
Click, 395, 111
Sleep 100
Click, 72, 751
Click, 72, 751
return


` & d::
hiv_comment := "��ĳHIV RNA �f�r�q�����Ч@�G���@����¦�A�Y�|���v���A��ĳ�C3-4�Ӥ�@�@���F�Y�}�l�v���A��ĳ�Ĥ@�Ӥ�n�@�@���A����3-4�Ӥ밵�@���A��ӥi4-6�Ӥ�@�@���C"
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