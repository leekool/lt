#Include <JSON>

#ifwinactive, ahk_class OpusApp,  ; IF WORD ACTIVE BEYOND THIS POINT

; capitalisation
:c:THe::The

; hard space
::mr::
sendinput	Mr^+{space}
input		key, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}
stringupper 	key, key ; uppercases next letter
sendinput	%key%
return

::ms::
sendinput 	Ms^+{space}
input, key, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}
stringupper, 	key, key
sendinput 	%key%
return

::mrs::
sendinput 	Mrs^+{space}
input, 		key, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}
stringupper, 	key, key
sendinput 	%key%
return

:c:para::
sendinput 	para^+{space}
return

:c:p::
sendinput 	p^+{space}
return

:c:subs::
sendinput 	subs^+{space}
return

:c:cl::
sendinput 	cl^+{space}
return

:c:et cetera::
sendinput 	et^+{space}cetera%A_EndChar%
return


; hard hyphen
:c:ce::
sendinput 	cross^+-examination%A_EndChar%
return

:c:ceing::
sendinput 	cross^+-examining%A_EndChar%
return

:c:ceed::
sendinput 	cross^+-examined%A_EndChar%

:c:CE::
sendinput 	CROSS^+-EXAMINATION%A_EndChar%
return

:c:oic::
sendinput 	officer^+-in^+-charge%A_EndChar%
return

:c:eic::
sendinput 	examination^+-in^+-chief%A_EndChar%
return

:c:takeaway::
sendinput 	take^+-away%A_EndChar%
return

; NT COMMENT (needs work)
:c:NT::
sendinput 	{shift}{backspace}..(not^+{space}transcribable)..^{left}^+{left 2}
sleep 		200
sendinput 	^!{m}								
keywait 	enter, D
sleep 		300
sendinput 	{esc}{end}  ; go back to main
sleep 		400
sendinput 	^!{:}  ; close comment pane
keywait 	+
sleep 		100
sendinput 	{shift up}{ctrl up}{alt up}  ; stop modifiers getting stuck (hopefully)
return

; break
^b::
sendinput 	{$}^+{- 2}{left 2}{backspace}{end}  ; bypasses word autocorrect
return

; send SPEAKER 1 - presiding
!1::
FileRead 	config, config.json
data := 	JSON.Load(config)
speaker1 := 	data.speaker1
stringupper, 	speaker1, speaker1
sendinput 	{end}{enter 2}%speaker1%:{space 2}
return

; send SPEAKER 2 - prosecutor/plaintiff
!q::
FileRead 	config, config.json
data := 	JSON.Load(config)
speaker2 := 	data.speaker2
stringupper, 	speaker2, speaker2
sendinput 	{end}{enter 2}%speaker2%:{space 2}
return

; send SPEAKER 3 - accused/defendant
!a::
FileRead 	config, config.json
data := 	JSON.Load(config)
speaker3 := 	data.speaker3
stringupper, 	speaker3, speaker3
sendinput 	{end}{enter 2}%speaker3%:{space 2}
return

; WITNESS
::TWW::  ; the witness withdrew
sendinput 	{<}THE WITNESS WITHDREW{enter 2}
return

; question
^q::  
keywait 	^				; fixes ctrl getting stuck (sometimes doesn't pls fix)
sendinput 	^+{left}^{insert}		; ctrl+insert is copy in Word
sleep 		100				; fix : not getting deleted
{									
if instr(clipboard, ":")
	sendinput {backspace}{enter 2}Q.{space 2}
else
	sendinput {right}{enter 2}Q.{space 2}
}
input, 		key, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}  ; capitalises next character
stringupper, 	key, key
sendinput 	%key%
return

; answer
^a::
sendinput 	{enter}A.{space 2}
input, 		key, L1, {LControl}{RControl}{LAlt}{RAlt}{LShift}{RShift}{LWin}{RWin}{AppsKey}{F1}{F2}{F3}{F4}{F5}{F6}{F7}{F8}{F9}{F10}{F11}{F12}{Left}{Right}{Up}{Down}{Home}{End}{PgUp}{PgDn}{Del}{Ins}{BS}{CapsLock}{NumLock}{PrintScreen}{Pause}
stringupper, 	key, key
sendinput 	%key%
return

; EXHIBIT
^+e::
inputbox, 	exhibit, exhibit number + exhibit
stringupper, 	exhibit, exhibit
sendinput 	EXHIBIT {#}%exhibit% TENDERED, ADMITTED WITHOUT OBJECTION{enter 2}
return
