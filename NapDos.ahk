#SingleInstance, Force
SendMode Input

Esc::ExitApp


G::

Path := "C:\Users\AndyN\Documents\Sevan-Andy Project\NapDos.xlsx"
ex := ComObjCreate("Excel.Application")
ex.visible := True
ex.Workbooks.Open(Path)
Sleep, 500
ex := ComObjActive("Excel.Application")

Sleep, 500

CoordMode, Mouse [, Window]

Start:

Run https://npdomainseeker.sdsc.edu/napdos2/run_analysis_v2.html
Sleep, 10000
MouseClick, Left, 508, 734, 1, 0
Sleep, 500
Send {PgDn}
Sleep, 500

; Choose File
MouseMove, 646, 866
Send {Click}
Sleep, 500
Send {Click}
Sleep, 500
Send {Delete}
Sleep, 2000
Send {Click}
Sleep, 180
Send {Click}
Sleep, 500

;Seek
MouseClick, Left, 525, 915, 1, 0
Sleep, 18000

;Submit Job
MouseClick, Left, 564, 473, 1, 0
Sleep, 18000

;Check if analysis is successful
PixelGetColor, pColor, 588, 677, RGB
if (pColor = 0xAABBDD)
    {
    MouseClick, Left, 588, 677, 1, 0
    Sleep, 3000
    MouseClick, Left, 578, 530, 1, 0
    Sleep, 3000
    Send {F6}
    Sleep, 500
    Send, ^c
    Sleep, 500
    ex.ActiveCell.value := clipboard ; data to excel active cell
    ex.ActiveCell.Offset(1,0).select ; select next cell where you paste next data
    Sleep, 20000
    Gosub, Carbon
    Return
    }

Else

    {
    Send {F6}
    Sleep, 500
    Send, ^c
    Sleep, 500
    ex.ActiveCell.value := clipboard ; data to excel active cell
    ex.ActiveCell.Offset(1,0).select ; select next cell where you paste next data
    Sleep, 20000
    Gosub, Carbon
    Return
    }

Carbon:
Send {Click}
Sleep, 10000
Send ^+{Tab}
Sleep, 500
Send !{NumpadLeft}
Sleep, 500
Send !{NumpadLeft}
Sleep, 500
Send !{NumpadLeft}
Sleep, 500
Send {PgUp}
Sleep, 500

MouseClick, Left, 510, 600, 1, 0
Sleep, 500
Send {PgDn}
Sleep, 500

;Seek
MouseClick, Left, 525, 915, 1, 0
Sleep, 15000

;Submit Job
MouseClick, Left, 564, 473, 1, 0
Sleep, 15000

;Check if analysis is successful
PixelGetColor, pColor, 588, 677, RGB
if (pColor = 0xAABBDD)
    {
    MouseClick, Left, 588, 677, 1, 0
    Sleep, 3000
    MouseClick, Left, 578, 530, 1, 0
    Sleep, 3000
    Send {F6}
    Sleep, 500
    Send, ^c
    Sleep, 500
    ex.ActiveCell.value := clipboard ; data to excel active cell
    ex.ActiveCell.Offset(1,0).select ; select next cell where you paste next data
    Sleep, 20000
    Gosub, Start
    Return
    }

Else

    {
    Send {F6}
    Sleep, 500
    Send, ^c
    Sleep, 500
    ex.ActiveCell.value := clipboard ; data to excel active cell
    ex.ActiveCell.Offset(1,0).select ; select next cell where you paste next data
    Sleep, 20000
    Gosub, Start
    Return
    }



;let GetResults=(URL) =>
;let
;    Source = Web.Page(Web.Contents(URL)),
;in GetResults

ExitApp