#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

MouseDown := 0

Esc::ExitApp

G::
Loop, 8

{
CoordMode, Mouse [, Window]
MouseClick, Left, 463, 632, 1, 0
Sleep, 180
MouseClick, Left, 593, 762, 1, 0
Sleep, 180
MouseMove, 1172, 769
Send {Click}
Sleep, 400
MouseMove, 0, %MouseDown%, 0, R
Sleep, 400
MouseDown += 26
Send {Click}
Sleep, 800
;Send {Delete}
;Sleep, 400
Send {Click}
Sleep, 400
Send {Click}
Sleep, 400
Send {PgDn}
Sleep, 180
MouseClick, Left, 332, 580, 1, 0
Sleep, 180
MouseClick, Left, 332, 635, 1, 0
Sleep, 180
MouseClick, Left, 1105, 726, 1, 0
Sleep, 180
Run https://antismash.secondarymetabolites.org/#!/start
Sleep, 7000
}

;Z::
;CoordMode, Mouse [, Screen]
;MouseClickDrag, Left,2309, 315, 390, 172 [, 10]
;MouseClick, Left, 2309, 315, 1, 0
;Sleep, 500
;Send ^x
;Sleep, 500
;MouseClick, Left, 390, 172, 1, 0

;Sleep, 500
;Send ^v
