#SingleInstance force

cnt_x = 332
cnt_y = 114

period_x = 43
period_y = 209


loop := 0
	
if not (WinExist("CTMdataStorage.xlsx - Excel"))
{
Run, CTMdataStorage.xlsx
}
if not (WinExist("Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote"))
{
MsgBox, "Error: Open Distribution Panel CCT @GfK and restart the process."
ExitApp
}
	
WinActivate, CTMdataStorage.xlsx
WinWait, CTMdataStorage.xlsx
Send, ^{up}
Sleep, 100
Send, ^{left}

Mainloop:
WinActivate, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
WinWait, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
if (loop <= 32)
	goto, ChangeCountry
else
{
	MsgBox, "Finished"
	ExitApp
}
return


ChangeCountry:
loop := loop+1
MouseClick, left, %cnt_x%, %cnt_y% ; CNT country dropdown
Sleep, 300
Send, {down}
Send, {enter}
; press Search
;MouseClick, left, %search_x%, %y1%
Sleep, 100
Send, {Tab 3}
Sleep, 100
Send, {Enter}
sleep, 1000
WinWaitClose, Loading... - \\Remote
sleep, 1000
If (WinExist("Distribution Panel CCT *ERROR - \\Remote"))
{
	WinActivate, Distribution Panel CCT *ERROR - \\Remote
	WinWait, Distribution Panel CCT *ERROR - \\Remote
	;MouseClick, left, 434, 139
	Send, {Enter}
}
WinActivate, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
WinWait, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
goto, CheckRecords
return


CheckRecords:
If (WinExist("TDistribution Panel CCT *WARNING - \\Remote"))
{
	WinActivate, TDistribution Panel CCT *WARNING - \\Remote
	WinWait, TDistribution Panel CCT *WARNING - \\Remote
	;MouseClick, left, 217, 140 ; press OK
	Send, {Enter}
	sleep, 1000
	WinActivate, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
	WinWait, Distribution Panel CCT @GfK - [dip.gfk.com][01.02.07] - \\Remote
	sleep, 500
	goto ChangeCountry
}
else
	goto Copydata

sleep, 1000
return


CopyData:
MouseClick, left, %period_x%, %period_y%
sleep, 1000
Send, ^a
sleep, 1000
Send, ^c
sleep, 1000
goto PasteInExcel
return

PasteInExcel:
WinActivate, CTMdataStorage.xlsx - Excel
WinWait, CTMdataStorage.xlsx - Excel
Send, ^v
sleep, 1000
Send, ^{Down}
sleep, 1000
Send, {Down}
goto Mainloop
return



Esc::ExitApp
