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

;Serial_Master working copy
#Requires AutoHotkey v2+
;msgbox("This program  is: AutoHotKey version " A_AhkVersion "`n line number " A_LineNumber)
Last_edit:= "HVP 09/30/2023" ; WinWaitClose in rx loop was nver exiting. Changed to new title and require exact match to exit.
; 09/19/2023" ; Program Options GUI was changing Globals on CANCEL.  Changes now done only on OK exit or init=true
;    WinExist(Partner_ID) is failing in RS232 COM port receive infinite loop after sleep
;    added  if !(WinWaitActive(Partner_ID, , Start_Delay)) { ; to give time to recover from wake-up
; 08/28/2023"  ;Added System UP Time to Serial_Master_V2_Splash
;   Added code to exit Screen Saver (if active) when serial input is received.
;   Note; Serial input resets Screen Saver timer only when the desktop is displayed.
; 08/20/2023"   ; Fixed errors when Partner_App="None" here in idle loop and
;   in Serial_Master_Menu_V2, Fixed Start_Program() & Part_Change(*) was causing Partner_App to change before OK button.
;   New_Partner_App was Partner_App.
; 08/11/2023   ;see Com Port Services, Com Config menu item now calls a top level function to open gui,
; then apply new settings.
; Changed HotkeySub() & HotkeySub_Ctrl_Enter() to end with Exit (not return) to terminate thread started by hotkey.
;07/23/2023
; changed serial transmit timer values disable Tx timeout.
; Fixed file start issues when file_name is open in another App.
; Now the 'official' Serial_Master.
;07/05/2023  now complete,
;   read loop
;   Read from RS232 ready for debug
;  RS_232 Write ready to test
;  RS232_Initialize() looks stable
;  Com_Config_GUI stable
;   Reads COM_Port.cfg and populates Gui settings
;   Writes COM_Port.cfg on gui close (OK), Closes, Opens, sets or updates COM port operating parameters.
;   Modem real time control lead display working.
;06/26/2023 Master_Menu_V2.ahk looks stable
Persistent     ; Keep the script running until the user exits it.
;#Warn All, Off
SetWorkingDir(A_ScriptDir)  ; makes the script starting directory the working directory.
#SingleInstance Force
SetTitleMatchMode(1)     ;A window's title must start with the specified WinTitle to be a match
global Version_Number:= "V2.0"
Global Partner_App       ; Partner Program name
Global Com_inifile:=  "COM_Port.cfg" ;name of the COM_Port  configuration file, must be in the same directory as this program file
Global Config_Message:=""   ; Message sent to a configuration Gui message area
Global Config_Font:=""      ; Config_Message font e.g. Config_Font:="s10 cRed" (size 10 color red)
Global RS232_Settings:=""   ; Serial port paeameters in 'mode' string format
                        ; mode syntax -- com<m>[:] [baud=<b>] [data=<d>] [parity=<p>] [stop=<s>] [xon={on|off}]
                        ;[odsr={on|off}] [octs={on|off}] [dtr={on|off|hs}] [rts={on|off|hs|tg}] [idsr={on|off}]
                        ; [to={on|off}] is not used. Timeouts are set in RS232_Initialize()
Global RS232_FileHandle:=0
Global RS232_Port:=""       ;
Global Splash_time:=4   ; Duration of splash screen
Global Key_List:=""     ; list of recognized decoder control keys

Global SPI_GETSCREENSAVERRUNNING := 0x0072
Global screen_saver_active:=0			;zero indicates the screen saver is not active
Global Start_Delay
;-----------------Install a custom tool tray Option list-----------------------------------------------------------------
Tray_modify:  ;Changes script menu to a custom list
TraySetIcon("Green_gear.png")	;image file must be in the working directory, where the script was started.

Tray:= A_TrayMenu               ;get the Tray menu object
;Tray.Delete() ; V1toV2: not 100% replacement of NoStandard,remove standard menu items
Tray.Add  ; Creates a separator line.
Tray.Add "Exit", Safe_Exit    ;
Tray.Add "Program Options", MenuHandler  ; Opens the Program Options Gui
Tray.Add "COM Port Configureation", Open_Com_Gui  ;Opens the COM configuration menu
Tray.Add("COM Port Configureation", Open_Com_Gui.bind(Com_inifile,0))  ;Opens the COM configuration menu
Tray.Add("About", Serial_Master_Splash.bind(0,Last_edit))  ; Creates the 'About' display.
Decoder_init()                              ; one time initiaize Key_List
Quit_var :=0                                ; Set by Ctrl-F1 hotkey to indicate intent to exit Serial Master
No_COM_Gui:=true
Com_Config_GUI(Com_inifile,No_COM_Gui)      ; get RS232_Settings and RS232_Port
MenuHandler(true) ;initialization only pass to initialize Global variables
Serial_Master_Splash(Splash_time,Last_edit) ; Open splash menu, continue to WinWaitClose()during count down
;########################################################################
;###### Build Hotkey Assignments -  Used to direct Keyboard input ######
; See https://learn.microsoft.com/en-us/windows/win32/inputdev/about-keyboard-input?redirectedfrom=MSDN
; for an overview of windows keyboard input process
;########################################################################

Loop 26      {    ;###### Direct Key Presses for a-z and Shift a-z
  Hotkey("$" chr(96+A_Index), HotkeySub)
  Hotkey("$+" chr(96+A_Index), HotkeySub)
}
Global Other_Keys:="``1234567890-=[]\;,./'" ;Does not include numeric keypad
N_Other_Keys:=StrLen(Other_Keys)
Loop N_Other_Keys {
  other_key:=SubStr(Other_Keys, (A_Index)<1 ? (A_Index)-1 : (A_Index), 1)
  Hotkey("$" other_key, HotkeySub)        ; Hotkeys Other_Keys...
  Hotkey("$+" other_key, HotkeySub)       ; and their shift
}
Hotkey("$Space", HotkeySub)
Hotkey("$ENTER", HotkeySub)
Hotkey("$BS", HotkeySub)
Hotkey("$esc", HotkeySub)
Hotkey("$tab", HotkeySub)
;########## Special Ctrl_Enter ########################################
  Hotkey("^enter", HotkeySub_Ctrl_Enter)  ;Enter sent only to Serial Port.
;########################################################################
^F1::               ;Ctrl F1 hotkey SCRAM
{
  Quit_var :=1         ;Kill Rx loop while MsgBox is open
  MsgBox("Serial Master is now disconnected from " RS232_Port "`n and will now exit`n Line number " A_LineNumber, "Serial Master " Version_Number)
  SAFE_EXIT()         ; Close com port, release all keys and Exit Script
  exit
}
;########################################################################
;######  END Hotkey Assignments                                 #########
;########################################################################

WinWaitClose("Serial Master " Version_Number " Welcome" )   ; Wait for Program_Splash GUI to close.
Start_Program()                                             ; Start the Partner program from Serial_Master.cfg
RS232_FileHandle:=RS232_Initialize(RS232_Settings)          ; Completes Com_port initialization and opens the port
RS232_Bytes_Received:=0
Loop {   ;RS232 COM port receive infinite loop
  if (Partner_App != "None") and !(WinExist(Partner_ID)) {  ; Partner Program = None, Data goes to App in focus...
    if !(WinWaitActive(Partner_ID, , Start_Delay)) {        ; WinExist may fail at sleep recovery,wait Start_Delay for wake up
        Config_Message:="Partner Program is not available to Serial Input data.`n Click OK to Re-Start " Partner_app
     . " or choose Partner Program None`n or EXIT to end Serial Master"
        Config_Font:="s10 cblue"
        MenuHandler()
        Save_Match_Mode:=A_TitleMatchMode
        SetTitleMatchMode(3)                          ; A window's title must match exactly to specified WinTitle to be a match.
        result:=WinWaitClose("Serial Master Options")     ; Wait for GUI to close.
        SetTitleMatchMode(Save_Match_Mode)
      }
  }
  Read_Data := RS232_Read(RS232_FileHandle,"0xFF",&RS232_Bytes_Received)
  ; Read_Data is a string where each received data byte is represented by 2 hex digit character code, values 00-FF may be included.
  ; 0xFF in the line above sets the size of the requested read buffer to 255 characters
  ; &RS232_Bytes_Received is a VarReference whose value is set by the called function to te number of data bytes received.
  ; Comm port time-outs are set to return immediately with any (including zero bytes) buffered receive data.
 ; Critical("On")                         ; Prevent interruption during execution of this loop.
  if (RS232_Bytes_Received > 0) {           ; Process the data, if there is any.
    ;MsgBox("RS232_FileHandle=" RS232_FileHandle "`n RS232_Bytes_Received=" RS232_Bytes_Received "`n Read_Data=" Read_Data "`n Line number " A_LineNumber)
;**************************************************************************************
; **** Serial input received.  Exit Screen Saver, if active.  *************************
;**************************************************************************************
  Dll_result := DllCall( "user32.dll\SystemParametersInfo"	; Working test for active screensaver
  , "uint", SPI_GETSCREENSAVERRUNNING
  , "uint", 0
  , "uint*", &screen_saver_active
  , "uint", 0 )
  if (Dll_result =0) {	; result =0 DLL failed
      msgbox( "SPI_GETSCREENSAVERRUNNING Dll failed `n" GetErrorString(A_LastError) , "Serial Master -- Line " A_LineNumber)
  }
  if (screen_saver_active != 0) {	; if !=0 ; Screensaver is active
    for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process") { ; find the screensaver (if any)
        if(InStr(process.Name, ".scr")){
            ProcessClose process.Name                               ; and close it.
            Break
        }
    }
    screen_saver_active:=0   ;Screen Saver should not be active
  }  ;******************************** End Screen saver exit *******************************
    if (Partner_App != "None") and !(WinActive(Partner_ID))  {   ; other Apps may become active when no receive data activity...
      ;if (WinExist(Partner_ID)) {          ; This really slows character through put!!
      WinActivate(Partner_ID)               ; Make sure the desired program gets this data
    }
    ;RS232_Read() returns Hex-ASCII data instead of an ASCII character bytes
    ASCII :=""                               ;Begin Data to ASCII character code conversion
    Loop RS232_Bytes_Received     {               ; loop for # input character codes. The decoder needs data one byte at a time
      Byte := SubStr(Read_Data, 1, 2)             ; Byte <- a 2 byte Hex-ASCII value
      Read_Data := SubStr(Read_Data, 3)           ; then remove 2 bytes from input data
      Byte := "0x" Byte                           ; Express as Hex format
      if (GIDEI_Decoder !="off")  {
        Byte:=Decode_serial(Byte)                 ; Decoder returns input it does not use or "" if used
      }
      if (Byte !="")       {                      ;V2 does not tolerate "" in Chr(Byte)
        ASCII .= Chr(Byte)                        ;Build ASCII string from available input
      }
    }      ;end of Loop
    SendText(ASCII)                      ;send all data received this READ pass as text characters to the in focus app
    if (Quit_var )
      return
  }                           ;end any input received
   ; Critical("Off")
}

;******************end receive infinite loop*****************************************
#include GIDEI_2_Decoder_v2.ahk
#include Com_Port_Services_V2.ahk
#include Serial_Master_Menu_V2.ahk
#include Serial_Master_V2_Splash.ahk
;####################################################################################################
;######  Keyboard data response function  Handles all QWERTY keys and their 'Shift'  ################
;################  All Hotkeys should have the '$' (use hook) prefix  ###############################
;####################################################################################################
;## Keyboard data is sent to the open RS232 COM port when Keymode is not 'Off'via
;## Serial Master V2 sends decimal character codes for SendInput() and ASCII-Hex strings to RS232_Write()
;##If Keymode = 'Off' No keyboard data is sent to the serial port, however, keyboard data is sent any in focus program.
;##If Keymode = "Serial" Keyboard data is sent ONLY the open RS232 COM port
;##          --- Note: If the serial port Echos receive data, this mode functions much like Force Keymode
;##IF Keymode = "Both" Keyboard data is sent to any in focus program and to the serial port.
;##IF Keymode = "Force", the Partner Application window is made active. Therefore, Keyboard data is sent to the serial port
;##              and Partner Application.

HotkeySub(ThisHotkey) {
Critical("on")
  if (A_ThisHotkey = "$Space")  {
    New_Key:=chr(32)
    }
  else if (A_ThisHotkey = "$BS")   {
      New_Key:=chr(8)
    }
  else if (A_ThisHotkey = "$Tab")   {
      New_Key:=chr(9)
  }
  else if (A_ThisHotkey = "$esc")   {
      New_Key:=chr(27)
  }
  else if (A_ThisHotkey = "$ENTER") {
      New_Key:=chr(13)                       ;Carriage Return (new line)
  }
  else {                                                ; only hotkeys for QWERTY keys and their Shift expected
    Global Other_Keys
    Other_Shift:="~!@#$%^&*()_+{}|:<>?" '"'             ; shift equivalent to Other_Keys
    New_Key:=StrReplace(A_ThisHotkey, "$")              ; Discardexpected $ (use hook) modifier
    New_Key:=StrReplace(A_ThisHotkey, "+", , , &Shift_Key)  ; Discard any +. Shift_Key:= # found
    New_Key:=SubStr(A_ThisHotkey, -1)                   ; Get the basic key pressed. ignore other modifiers
    if (Shift_Key)  {                                   ; Was there a + in A_ThisHotkey?
      if(FoundPos:=(InStr(Other_Keys, New_Key)))  {     ; Is an irregular key shift?
        New_Key:=SubStr(Other_Shift, FoundPos, 1)       ; make the substitution
      }
    }
    if ((GetKeyState("CapsLock","T") ^ Shift_Key )) {
      New_Key := StrUpper(New_Key)                      ; make only a-z upper case
    }
  }
  if (Key_Mode ="Force") and  (Partner_App != "None")   {
    WinActivate(Partner_ID)                             ; Force the data to Partner_App
  }
  if (Key_Mode !="Serial")     {                        ;"Serial" Suppresses keystroks to local App
    SendInput("{Text}"  New_Key)
  }
  if (Key_Mode !="off") and (Quit_var =0) {             ; Data to serial port?
    New_Key :="0x" Format("{:X}", Ord(New_Key))         ; Get the Hex-ASCII equivalent
    if(StrLen(New_Key) =3)  {
      New_Key:= StrReplace(New_Key,"0x","0x0")          ; make this a proper Hex-ASCII Value
    }
    if (New_Key ="0x0D") {
      New_Key .=",0x0A"                                 ;add line feed to carrage return
    }
    RS232_Write(RS232_FileHandle,New_Key)               ;Send it out the RS232 COM port
    ;MsgBox(  New_Key " Sent to serial port`n Line number " A_LineNumber)
  }
Critical("off")
  return

}
;########## Special Ctrl_Enter ########################################
HotkeySub_Ctrl_Enter(ThisHotkey) {
  if (Key_Mode !="off") {
    var:="0x0D,0x0A"                                    ;Carriage Return,New Line
    RS232_Write(RS232_FileHandle,var)                   ;Send it out the RS232 COM port
  }
  if (Key_Mode ="Focus") {
    if not WinActive(Partner_ID) {
      WinActivate(Partner_ID)                           ;make sure the desired program gets this data
    }
  }
  if !(Key_Mode ="Serial") {
   SendInput(A_Space)                                   ;and a space to the App
  }
  ;msgbox("Ctrl-Enter pressed...")
return
}
;########################################################################
;###### End of Serial Port Transmit and keyboard data handler ###########
;########################################################################


Safe_Exit(*)	{					; see Varadic Functions for the meanimg of (*)
  RS232_Close(RS232_FileHandle)
  if (GIDEI_Decoder !="off")  {
    Send_esc("<esc>,rel.")
    Send_esc("<esc>,mourel.")
  }
  ExitApp()
}

;*********************************************************************************************************************************
;********************************************* System wide utility functions  ****************************************************
;*********************************************************************************************************************************

Config_Default_Msg(file, item, default_Val, Ln_Num:="Unknown" ) {
  MsgBox("No valid " item " , default to " default_Val " selected`n Line number " Ln_Num,  file "  File error!" )
  return
}

GetErrorString(Errornumber)             ;Usually set to A_LastError
  {
	VarSetStrCapacity(&ErrorString, 1024)		;String to hold the error-message.
	DllCall("FormatMessage"         ;FORMAT_MESSAGE_FROM_SYSTEM:
            , "UINT",  0x00001000	;The function should search the system message-table resource(s) for the requested message.
			, "UINT", 0             ;A handle to the module that contains the message table to search.
			, "UINT", Errornumber
			, "UINT", 0				;Language-ID is automatically retreived
			, "Str" , ErrorString
			, "UINT", 1024			;Buffer-Length
			, "str" , ""
            )
	ErrorString := StrReplace(ErrorString, "`r`n", A_Space)		;Replaces newlines by A_Space for one line-output
	return ErrorString
  }