/* Copyright 2023 Harry V. Paul
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
 * files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy,
 * modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software
 * is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
 * OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
 * CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
#Requires Autohotkey v2.0+
;msgbox(" This program  is: AutoHotKey version " A_AhkVersion "`n line number " A_LineNumber)
;################### Serial_Master_V2_Splash.ahk ###################
; Last_edit:= "HVP 08/28/2023"  :Added System UP Time
; 08/04/2023"
; Splash_Timer() & Button_Splash_Done(*) end in Exit (not return) to end thread.
;06/28/2023
;Splash_time = 0: About display with OK button to close. Otherwise, Countdown splash lasting Splash_time Seconds
;#################################################################################################################
;#################  Timed Splah and Button Close Splash Screens    ###############################################
;#################################################################################################################
;global Partner_App         	; (Read only) Partner Program name
;global RS232_settings			; (Read only) RS-232 Mode settings
;global Version_Number 			; (Read only) Version number of this program
Serial_Master_Splash(Splash_time,Last_edit, *)	{		;entry point for a count-down or 'about' splash screen

	TimeFormat := "dddd	HH:mm:ss"
	UpTime := A_TickCount // 1000         ; Elapsed seconds since start
	;StartTime:= DateAdd(A_Now, -UpTime, "seconds")       ;
	;StartTime:=FormatTime( StartTime, TimeFormat)
	;NowTime:=FormatTime(A_Now , TimeFormat)  ;
	uptime_min:=(StrLen(mod(UpTime // 60, 60)) <2 ) ? "0" mod(UpTime // 60, 60) : mod(UpTime // 60, 60)
	uptime_sec:= (StrLen(mod(UpTime, 60)) < 2) ? "0" mod(UpTime, 60) : mod(UpTime, 60)
	UpTime := "System UP Time   " UpTime // 86400 " days " mod(UpTime // 3600, 24) ":" uptime_min ":" uptime_sec
	;MsgBox "System Up Time:`t" UpTime "`n`nStart time `t" StartTime "`nTime now:`t" NowTime , ,64
Splash_height :=300				;Sets the vertical size of the box
Program_Splash := Gui(,"Serial Master "  Version_Number " Welcome")
Program_Splash.Opt("-SysMenu")	    ; removes icon and control buttons from title bar
Program_Splash.Opt("+alwaysontop")	;another window will not hide this pop-up
Program_Splash.MarginX:= 10       	;replaces Program_options.Margin ("10", "10")from V1
Program_Splash.MarginY:= 10
Program_Splash.SetFont "s14 bold", "Times New Roman"
Program_Splash.Add("Text", , "Serial Master integrates a Serial COM_Port with a Partner Program")
{ ; This forms a new single line with two colors of text (braces not required)
Program_Splash.SetFont "cblack"
Program_Splash.Add("Text", "xm+5", "Right click the")
Program_Splash.SetFont "cgreen"
Program_Splash.Add("Text", "x+5", "green gear")
Program_Splash.SetFont "cblack"
Program_Splash.Add("Text", "x+5", "icon in the tool tray, then left click a menu item'")
}
Program_Splash.Add("Text", "xm+10 y+5", "'About' to view this message.")
Program_Splash.Add("Text", "xm+10 y+5", "'COM Port Configuration' to change COM Port number or set advanced parameters")
Program_Splash.Add("Text", "xm+10 y+5", "'Program Options' to set Partner Program Parameters.")
Program_Splash.SetFont "s14 norm cDA4F49"   ;no longer bold, a soft shade of red
Program_Splash.Add("Text", "xm", Last_edit)	;position, same x, y (y= standard new line + y last line
Program_Splash.SetFont "s10 cblack"
Program_Splash.Add("Text", "x+40", UpTime)	;position,
Program_Splash.SetFont "s14 cBlue"
Program_Splash.Add("Text", "xm y+20", "Serial Master includes a GIDIE Standard Escape Sequence Decoder.")	;text at x margin, down 20 pix
Program_Splash.SetFont "s14 norm"
if (Splash_time =0) {					  	; OK Button splash
	Program_Splash.Add("Text", "xm y+20", Partner_App " is the active Partner Program"	)
	Program_Splash.SetFont "s11 cblack norm", "Verdana"
	; add an OK button and goto 'Button_Splash_Done' when clicked or ENTER
	Program_Splash.Add("Button", "x10 y" Splash_height -35  " default", "OK").OnEvent("click", Button_Splash_Done)
	Program_Splash.Add("Text", "x+5 y" Splash_height -25, "Click to close")
	}
else {								    ;count down splash
	Splash_hint := "This window will close in " Splash_time " seconds."
	Program_Splash.Add("Text", "xm y+20", "Preparing to open " Partner_App)
	Program_Splash.SetFont "s11 cblack norm", "Verdana"
	TextSplash_end:=Program_Splash.Add("Text", "xm y" Splash_height -25 " vSplash_end", Splash_hint)			;
	SetTimer(Splash_Timer,1000)	    ; set a timer for 1 second tic interval. Splash_Timer; counts down the splash screen
  }
Program_Splash.SetFont "s10 cblue"
Program_Splash.Add("Text", "x+10", RS232_settings)
Program_Splash.Show("w800 h" Splash_height)
return		;still more work to be done on this thread
	Splash_Timer()							;Used in conjunction with a Splash screen with a count down timer
	{
		--Splash_time
		Splash_hint := StrReplace(Splash_hint, Splash_time+1, Splash_time)	;replace the 'old' time with 'new'.
		if (Splash_time =1)  {
		  Splash_hint := StrReplace(Splash_hint, "seconds", "second")
		}
		TextSplash_end.Value := Splash_hint		;replace Gui text with Splash_hint as modified
		if (Splash_time = 0)  {					;last second expired?
		  SetTimer(Splash_Timer,0)				; delete timer
		  Program_Splash.Destroy()
		}
		Exit                    ; end timer thread
	}
	Button_Splash_Done(*)            		;arrive here by Program_Splash OK button
	{
		Program_Splash.Destroy()
		Exit                    ; end this thread
	}
}

;*************************END Splash screen Components   ************************************************