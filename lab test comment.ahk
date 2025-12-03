;#Requires AutoHotkey v1.0
;Please open and save with Big5

interrupted := false   ; default: not interrupted
Esc::
interrupted := true
return


` & space::
Click, 395, 111
Sleep 100
Click, 72, 751
Click, 72, 751
return


` & v::
InputBox, sp_number, Verficator, Please enter the specimen numbers:,
Loop, %sp_number%
{
    Click, 395, 111
    Sleep 100
    Click, 72, 220
    Click, 72, 220
}
return

` & d::
hiv_comment := "This test utilizes the Roche 2.0 assay to measure HIV RNA viral load, with a limit of quantitation as few as 20 copies per milliliter. It is recommended that HIV RNA viral load be measured twice as a baseline. If the patient has not yet started treatment, testing should be performed every 3–4 months. If treatment has begun, testing is advised at one month after initiation, followed by every 3–4 months, and subsequently every 4–6 months as needed."

Clipboard := hiv_comment
InputBox, sp_number, L72-999 commenter, Please enter the specimen numbers:,
Loop, %sp_number%
{
    Click, 1295, 716 ;select the comment field
    Send ^v ; send HIV comment
    Sleep 100
    Click, 395, 111 ; verify the specimen
    Sleep 100
    Click, 73, 220
    Click, 73, 220 ;double click the button of next specimen
}
return

` & r::
rpr_comment := "For syphillis screening cases, reactive RPR results indicate TPPA or FTA-ABS for confirmation, nonreactive RPR results preliminarily exclude syphilis, but follow-up is recommended in highly suspicious cases. For cases with syphillis history, RPR titer reduction more than 4 folds indicates succuessful treatment."
rpr_screening_comment := "In syphillis screening settings, reactive RPR results indicate TPPA or FTA-ABS for confirmation, nonreactive RPR results preliminarily exclude syphilis, but false negative could be possible during window period. Suggest rechecking in highly suspicious cases."
vdrl_comment := "The diagnosis of neurosyphilis requires a reactive serum treponemal test (FTA-ABS, TPPA, etc). In patients with neurologic symptoms, reactive CSF VDRL indicates neurosyphilis treatment; but nonreactive CSF VDRL does NOT rule out neurosyphilis. For patients with nonreactive CSF VDRL, suggest checking CSF WBC, CSF protein(cases without HIV), and CSF FTA-ABS(cases with HIV)."
rpr_newborn_comment := "Lab diagnosis of congenital syphilis requires an infant RPR/VDRL titer ≥4-fold the birthing parent’s titer. If this is not met, clinical and maternal histories should be reviewed. Do not order TPPA or FTA-ABS unless parental results are unknown, because the antibodies detected are of maternal origin."
rpr_comment_safety_testing := "For syphillis screening cases, reactive RPR results indicate TPPA or FTA-ABS for confirmation, nonreactive RPR results preliminarily exclude syphilis, but follow-up is recommended in highly suspicious cases. For cases with syphillis history, RPR titer reduction more than 4 folds indicates succuessful treatment.."

InputBox, sp_number, L72-101 commenter, Please enter the specimen numbers:,
Loop, %sp_number%
{
    if (interrupted){
        MsgBox, Script has been stopped by user.
        break
    }
    Clipboard := "" ; empty the clipboard for next interpretation
    ClipWait, 1 ; waiting clipboard to change in a max 5 seconds
    Click, 1295, 716
    Click, 1295, 716 ; select the comment field
    Send ^c
    ClipWait, 1 ; waiting clipboard to change in a max 1 seconds
    if (Clipboard = ""){ ; if empty, this means there is no comment in the comment field.
        Click, 1027, 170
        Click, 1027, 170 ; select the age
        Send ^c
        ClipWait, 1
        num := Clipboard + 0  ; turn the age into float type
        if (num > 77 ){
            Clipboard := rpr_newborn_comment
            ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
            Click, 1295, 716 ;select the comment field
            Send ^v ; send comment
        }
        else{
            Click, 888, 203
            Click, 888, 203 ; select the specimen
            Send ^c
            ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
            if (Clipboard = "B"){
                Click, 311, 203
                Click, 311, 203 ; select the patient origin
                Send ^c
                ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
                if (Clipboard = "H"){
                    Clipboard := rpr_screening_comment
                    ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
                    Click, 1295, 716 ;select the comment field
                    Send ^v ; send comment
                }
                else{
                    Clipboard := rpr_comment
                    ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
                    Click, 1295, 716 ;select the comment field
                    Send ^v ; send comment
                }

            }
            else if(Clipboard = "CSF"){
                Clipboard := vdrl_comment
                ClipWait, 5 ; waiting clipboard to change in a max 5 seconds
                Click, 1295, 716 ;select the comment field
                Send ^v ; send comment
            }
        }
    }
    Sleep 100
    Click, 395, 111 ; verify the specimen
    Sleep 100
    Click, 73, 220
    Click, 73, 220 ;double click the button of next specimen

}
return