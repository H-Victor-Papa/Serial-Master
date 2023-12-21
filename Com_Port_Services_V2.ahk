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
#Requires Autohotkey v2.0.9+
; last edit 12/13/2023 Added Com_Port_Init(inifile,No_Gui:=false). Create config file,
;	if No_Gui is false: open Gui for editing, then open port.
; 09/19/2023 "Recursion limit error" on ExitApp fixed in v2.0.9. Original code restored.
;	Active MsgBox statements put "Com Port Services Line number " A_LineNumber in title
; 08/11/2023 fixed circular reference caused by Com_Config_GUI() & RS232_Initialize().
; 08/09/2023  --work arround for issue in RS232_Read() that caused "Recursion limit error" on ExitApp
; 08/04/2023 -- Modem_Leads() ends in Exit (not return) to end thread.
; 07/28/2023  -- Tx timouts are now off, DCB now complete for Xon-xoff
;###############Com_Port_Services_V2.ahk###################################################################
; This is a code module that when #include (ed) in the Serial Master program provides all functions needed
; to comfigure and operate a Serial Comm Port. The functions have built-in error recovery processes that
; require little to no external support code. This code is derived from Serial Master V1 to V2 conversion
; Read and Write operations use synchronous or non-overlapped I/O
;
; Included functions:
;	A COM_Port.cfg file stores current port configuration data in the directory of this module.
;	Open_Com_Gui(*)   Calls Com_Config_GUI(com_inifile) and applies resulting settings when Gui is closed.
;   Com_Config_GUI(inifile [,No_Gui],*)    No_Gui is optional, default =false
;		Populates RS232_Settings (a mode string) from inifile and returns if No_Gui =true.  Else,
;		Opens a serial port configuration Gui that when closed by clicking 'OK', saves configuration
;		data to COM_Port.cfg, populates RS232_Settings and , if the Gui changes the COM port, the 'Old"
;		Com_port is closed and Serial_Master Menu is called to open a new file if necessary.
;		Configuration changes are NOT applied to the port.
;		If the Gui is closed by clicking 'CANCEL' the gui is closed without change to the COM port status or COM_Port.cfg.
;		The Gui provides a real time display of DCD, DSR, CTS and RI Modem interface controls of an open
;		COM port.
;
;	RS232_Close(RS232_FileHandle)  		Closes the COM port 'RS232_FileHandle'
;	RS232_Initialize(RS232_Settings)	Opens a Com port with parameters in RS232_Settings. Returns RS232_FileHandle
;										Note that Com_Config_GUI() may be called if the requested port fails to open.
;	RS232_Write(RS232_FileHandle,Tx_Message [,Data_Type])  Data_Type is optional, default=1
;		RS232_FileHandle,  FileHandle to an Open RS-232 Com port
;		Tx_Message,  A string containing the data to be sent.  Data_Type identifies the format of string content.
;		Data_Type,	If equal to 1, H or h Tx_Message is a comma delimited string of ASCII character code numbers.
;					ASCII character code numbers may be decimal (0-255) or hex ("0x00"-"0xFF") e.g Tx_Message;="0x0A,0x0D".
;					All other Data_Type strings are treated as a 'normal' Null Terminated String.
;		Returns: Number of Bytes_Sent
;	RS232_Read(RS232_FileHandle, Num_Bytes, &RS232_Bytes_Received)
;		RS232_FileHandle,  FileHandle to an Open RS-232 Com port
;		Num_Bytes,	The Maximum number of receive bytes this call will return.
;		&RS232_Bytes_Received,  Variable is updated with the actual number of received bytes returned when
;								all available receive bytes have been returned.
;		Returns:	A string of 2 byte Hex-ASCII character codes. Each character code may be
;					("0x00"-"0xFF") allowing any 7 or 8 bit ASCII code to be returned, including NULL.
;
;##########################################################################################################
;################ GUI control window to gather and groom COM Port configuration parameters  ###############
;##################   Reads inifile  (in this program directory)               ############################
;################## Updates inifile on 'OK' exit                               ############################
;##################   and returns a mode string  in RS232_Settings ########################################
;##########################################################################################################
;~ global Config_Message :=""
;~ Global config_font:=""
;~ Global RS232_FileHandle:=0
;~ inifile:="COM_Port.cfg"
;~ Com_Config_GUI(inifile)
;~ return
;~ Config_Default_Msg(file, item, default_Val, Ln_Num:="Unknown" ) {
  ;~ MsgBox("No valid " item " , default to " default_Val " selected`n Line number " Ln_Num,  file "  File error!" )
  ;~ return
;~ }

Com_Port_Init(inifile,No_Gui:=false) {
Global Config_Message
Global config_font
;Create config file, Open GUI with message and wait for close
FileAppend "[Default]`n"
   . "COM_Port_N=COM3`n"
   . "Baud_Rate=57600`n"
   . "N_Data=8`n"
   . "Parity_bit=None`n"
   . "Stop_Bits=1`n"
   . "DTR_State=Off`n"
   . "RTS_State=On`n"
   . "Tx_Flow_Control=None`n"
   . "Rx_Flow_Control=None`n`n"
   . "[COM_Configuration]`n"
   . "COM_Port_N=COM4`n"
   . "Baud_Rate=57600`n"
   . "N_Data=8`n"
   . "Parity_bit=None`n"
   . "Stop_Bits=1`n"
   . "Tx_Flow_Control=None`n"
   . "Rx_Flow_Control=None`n"
   . "DTR_State=On`n"
   . "RTS_State=On`n"
, inifile
	if (No_Gui) {
		return
		}
	Config_Font:="s10 cRed"
    Config_Message:="Default Configuration file Created"
	Com_Config_GUI(inifile,false,false)	;Open gui, with message, for editing. No resync with Master_Menu
	WinWaitClose("Com Port Configuration")  ; Wait for GUI to close.
	return
}


Open_Com_Gui(*)	{						; Menu selection lands here
	;global Com_inifile					; Com_inifile is global (read only)
	Com_Config_GUI(Com_inifile)
	WinWaitClose("Com Port Configuration")  ; Wait for GUI to close.
	RS232_Initialize(RS232_Settings)       	; Apply new settings
return
}
Com_Config_GUI(inifile,No_Gui:=false,X_Sync:=true,*) {
{
;The following variables are Global in scope and may be change herein.
Global Config_Message
Global config_font
Global RS232_Settings
Global RS232_Port			;May be used in Master Menu as part of a file_name
Global RS232_FileHandle

; The following variables are Global in scope and 'READ ONLY'herein.


static NoHide:=0
static Rate_Arr :=["110","300","600","1200","2400","4800","9600","19200","38400","57600","115200"]		; recognized baud rates
static N_bits_Arr :=["7","8"]												; number of data bits supported
static P_bit_Arr :=["Odd","Even","None"]		                          	; recognized parity bit choices
static Tx_Flow_Arr :=["XON/XOFF","CTS","DSR","All","None"]			       	; recognized Tx_Flow choices
static Rx_Flow_Arr :=["XON/XOFF","RTR (RTS)","DTR","DSR","None"]		  		; recognized Rx_Flow choices
static DTR_Arr :=["Off","On","DTR_Rx_Flow*"]		                      	; recognized DTR choices
static RTS_Arr :=["Off","On","Data_Rdy","RTR_Flow*"]		               	; recognized RTS choices
static COM_Arr :=["COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8"]	; This program can accept com ports with 2 digits


if (RS232_FileHandle >0)  {
	SetTimer(Modem_Leads,100)            ; start a 100mS timer to sample Modem leads
}
;########### Get port parameters from an ini.cfg file for the GUI ######################################################
; inifile 				Points to COM_Port.cfg.
	COM_Port := IniRead(inifile, "COM_Configuration", "COM_Port_N","ERROR")
	if (COM_Port= "ERROR" or COM_Port= "" ) { ; Pop message, use default
		COM_Port := IniRead(inifile, "Default", "COM_Port_N","ERROR")				;if this is ERROR, the config file should be presumed missing
		Config_Default_Msg(inifile, "COM port", COM_Port, A_LineNumber)
	}
	COM_num := SubStr(COM_Port, 4, 1)	;Keep only the port number for the drop down Choose.
	Old_Port:=COM_Port
	Bit_Rate := IniRead(inifile, "COM_Configuration", "Baud_Rate","ERROR")
	if (Bit_Rate = "ERROR" or Bit_Rate ="" ) {  ; Pop message, use default
		Bit_Rate := IniRead(inifile, "Default", "Baud_Rate","ERROR")
		Config_Default_Msg(inifile, "Baud_Rate", Bit_Rate,A_LineNumber)
	}
	Bit_Rate_Pick:=InArr(Rate_Arr, Bit_Rate)							; A numbr is needed here

	N_bits := IniRead(inifile, "COM_Configuration", "N_Data","ERROR")				;Read the number of daya bits
	if (N_bits = "ERROR" or N_bits ="") { ; Pop message, use default should also check range
		N_bits := IniRead(inifile, "Default", "N_Data","ERROR")
		Config_Default_Msg(inifile,"N_Data", N_bits,A_LineNumber)
	}
	N_bits_Pick:=InArr(N_bits_Arr, N_bits)
	P_bit := IniRead(inifile, "COM_Configuration", "Parity_bit","ERROR")				;Read the parity bit type
	if (P_bit = "ERROR" or P_bit = "")  { ;Pop message, use default
		P_bit := IniRead(inifile, "Default", "Parity_bit","ERROR")
		Config_Default_Msg(inifile, "Parity_bit", P_bit,A_LineNumber)
	}
	P_bit_Pick:=InArr(P_bit_Arr, P_bit)

	S_bits := IniRead(inifile, "COM_Configuration", "Stop_Bits","ERROR")
	if (S_bits= "ERROR" or S_bits = "") { ; Pop message, use default
		 S_bits := IniRead(inifile, "Default", "Stop_Bits","ERROR")
		 Config_Default_Msg(inifile, "Stop_Bits", S_bits,A_LineNumber)
	}

	Tx_Flow := IniRead(inifile, "COM_Configuration", "Tx_Flow_Control","ERROR")
	if (Tx_Flow ="ERROR" or Tx_Flow = "") { ; Pop message, use default
		Tx_Flow := IniRead(inifile, "Default", "Tx_Flow_Control","ERROR")
		Config_Default_Msg(inifile, "Tx_Flow_Control", Tx_Flow,A_LineNumber)
	}
	Tx_Flow_Pick:=InArr(Tx_Flow_Arr, Tx_Flow)
	Rx_Flow := IniRead(inifile, "COM_Configuration", "Rx_Flow_Control","ERROR")
	if (Rx_Flow= "ERROR" or Rx_Flow ="") { ; Pop message, use default
		Rx_Flow := IniRead(inifile, "Default", "Rx_Flow_Control","ERROR")
		Config_Default_Msg(inifile, "Rx_Flow_Control", Rx_Flow,A_LineNumber)
	}
	Rx_Flow_Pick:=InArr(Rx_Flow_Arr, Rx_Flow)

	DTR_bit := IniRead(inifile, "COM_Configuration", "DTR_State","ERROR")
	if ((DTR_bit = "ERROR") or (DTR_bit=""))  ; Pop message, use default
	{
		DTR_bit := IniRead(inifile, "Default", "DTR_State","ERROR")
		Config_Default_Msg(inifile, "DTR_State", DTR_bit,A_LineNumber)
	}
	DTR_Pick:=InArr(DTR_Arr, DTR_bit)
	RTS_bit := IniRead(inifile, "COM_Configuration", "RTS_State","ERROR")
		if (RTS_bit= "ERROR" or RTS_bit="") { ; Pop message, use default
		RTS_bit := IniRead(inifile, "Default", "RTS_State","ERROR")
	Config_Default_Msg(inifile, "RTS_State", RTS_bit,A_LineNumber)
	}
	RTS_Pick:=InArr(RTS_Arr, RTS_bit)

	if(No_Gui=true)  {  				; Build simple RS232_Settings and return
		RS232_Port:=COM_Port			; and set RS232_Port
		RS232_Settings:=(COM_Port
			":baud=" Bit_Rate
			" data=" N_bits
			" parity=" SubStr(P_bit, 1, 1)
			" stop=" S_bits
			" dtr=" DTR_bit
			" rts=" RTS_bit)
		;MsgBox("Simple mode string" RS232_Settings)
		return
	}


; any reference to controls in this GUI that are outside the thread of this GUI must reference the name as part of the
;first parameter e.g.  GuiControl,Com_cfg_Gui: , DSR_State,100.

Com_cfg_Gui := Gui(,"Com Port Configuration")
Com_cfg_Gui.OnEvent("Close", Com_cfg_GuiButtonOK)
Com_cfg_Gui.OnEvent("Escape", Com_cfg_GuiButtonCANCEL)
Com_cfg_Gui.MarginX:= 10			;replaces Com_cfg_Gui.Margin("10", "10")
Com_cfg_Gui.Marginy:= 10
Com_cfg_Gui.Add("Button", "x220 y10 default", "OK").OnEvent("Click", Com_cfg_GuiButtonOK)  ;
Com_cfg_Gui.Add("Button", "x+5 y10", "CANCEL").OnEvent("Click", Com_cfg_GuiButtonCANCEL)
Com_cfg_Gui.Add("GroupBox", "x6 y10 w80 h40", "Ports")
Com_cfg_Gui.Add("DropDownList", "x15 y25 w60 vCOM_Port_out Choose" COM_num, COM_Arr)
Com_cfg_Gui.Add("GroupBox", "x+25 y10 w80 h40", "Bit Rate")
Com_cfg_Gui.Add("DropDownList", "x110 y25 w60 vBit_Rate_out Choose" Bit_Rate_Pick, Rate_Arr)
Com_cfg_Gui.Add("GroupBox", "x6 y55 w300 h65", "Character format")
Com_cfg_Gui.Add("GroupBox", "x6 y70 w80 h40", "Data bits")
Com_cfg_Gui.Add("DropDownList", "x15 y85 w50 vN_bits_out Choose" N_bits_Pick, N_bits_Arr)
Com_cfg_Gui.Add("GroupBox", "x100 y70 w75 h40", "Parity bit")
Com_cfg_Gui.Add("DropDownList", "x110 y85 w60 vP_bit_out Choose" P_bit_Pick, P_bit_Arr)
Com_cfg_Gui.Add("GroupBox", "x190 y70 w75 h40", "Stop Bits")
Com_cfg_Gui.Add("DropDownList", "x200 y85 w50 vS_bits_out Choose" S_bits, ["1","2"])
Com_cfg_Gui.Add("GroupBox", "x6 y120 w100 h40", "Tx Flow Control")
Com_cfg_Gui.Add("DropDownList", "x15 y135 w80 vTx_Flow_out Choose" Tx_Flow_Pick, Tx_Flow_Arr)
Com_cfg_Gui.Add("GroupBox", "x120 y120 w100 h40", "Rx Flow Control")
RX_Flow_opt:=Com_cfg_Gui.Add("DropDownList", "x126 y135 w80 vRx_Flow_out  Choose" Rx_Flow_Pick, Rx_Flow_Arr)
RX_Flow_opt.OnEvent("Change", New_Rx_Flow)   			;make adjustments when this parameter changes...
Com_cfg_Gui.Add("GroupBox", "x6 y165 w100 h40", "RTS Control")
RTS_Opt_DDL:=Com_cfg_Gui.Add("DropDownList", "x20 y180 w80 vRTS_out Choose" RTS_Pick, RTS_Arr)
RTS_Opt_DDL.OnEvent("Change", New_Rts)
Com_cfg_Gui.Add("GroupBox", "x120 y165 w100 h40", "DTR Control")
DTR_Opt_DDL:=Com_cfg_Gui.Add("DropDownList", "x125 y180 w80 vDTR_out Choose" DTR_Pick, DTR_Arr)
Com_cfg_Gui.Add("GroupBox", "x6 y205 w300 h40", "Message Area")
Com_cfg_Gui.SetFont config_font
Msg_Var:=Com_cfg_Gui.Add("text", "vConfig_Msg_Var x10 y225 w300 h20", Config_Message )
Com_cfg_Gui.SetFont()
Com_cfg_Gui.Add("GroupBox", "x6 y250 w300 h65", "Control leads In")
Com_cfg_Gui.Add("text", "x45 y270 w300 h20", "DCD                DSR                CTS                RI")
DCD_Lead_Disp:=Com_cfg_Gui.Add("Progress", "x26 y270 w15 h15 cgray vDCD_State", "100")
DSR_Lead_Disp:=Com_cfg_Gui.Add("Progress", "x96 y270 w15 h15 cgray vDSR_State", "100")
CTS_Lead_Disp:=Com_cfg_Gui.Add("Progress", "x166 y270 w15 h15 cgray vCTS_State", "100")
RI_Lead_Disp:=Com_cfg_Gui.Add("Progress", "x236 y270 w15 h15 cgray vRI_State", "100")
Com_cfg_Gui.BackColor :="F0F0F0"
Com_cfg_Gui.Show("x232 y138 h297 w316")
Com_cfg_Gui.Opt("+alwaysontop") 		;another window will not hide this pop-up
Com_cfg_Gui.Opt("-SysMenu")			; removes icon and control buttons from title bar
Com_cfg_Gui.Submit(NoHide)  		; Save each control's contents to its associated variable.
Config_Message:=""
  ;~ MsgBox(" COM port = " COM_Port "`n Bit_Rate = " Bit_Rate "`n databits = " N_bits "`n Parity bits = " P_bit
  ;~ . "`n stop bits = " S_bits "`n Tx_Flow_Control = " Tx_Flow  "`n Rx_Flow_Control = "  Rx_Flow "`n DTR_State = " DTR_bit
  ;~ . "`n RTS_State = " RTS_bit "`n Line number " A_LineNumber," Configuration Values")
return	;Keeps Gui thread open

Com_cfg_GuiButtonOK(*) {
	Options_result := Com_cfg_Gui.Submit(NoHide)   ; Save each control's contents to its Options_result variable.
	;----------------------------Update Configuration file with current COM variables-------------------------------
	;IniWrite, 'Pairs', Filename, Section requires that ALL parameter 'Pairs' of a section must be supplied...
	; Note that 'new_line' separates list 'Pairs'...
	;-----------------------------------------------------------------------------------------------------------
	IniWrite("COM_Port_N=" Options_result.COM_Port_out
	"`nBaud_Rate=" Options_result.Bit_Rate_out
	"`nN_Data=" Options_result.N_bits_out
	"`nParity_bit=" Options_result.P_bit_out
	"`nStop_Bits=" Options_result.S_bits_out
	"`nTx_Flow_Control=" Options_result.Tx_Flow_out
	"`nRx_Flow_Control=" Options_result.Rx_Flow_out
	"`nDTR_State=" Options_result.DTR_out
	"`nRTS_State=" Options_result.RTS_out, inifile, "COM_Configuration")
	;-----------------------------Configuration file Update done-----------------------------------------------------------------
	; Now assemble the components of a Mode string that may require adjustment. All others are taken directly from Options_result
	P_bit := SubStr(Options_result.P_bit_out, 1, 1)	;Keep only the first letter of config choice
	if (Options_result.Rx_Flow_out="XON/XOFF") or (Options_result.Tx_Flow_out="XON/XOFF") or (Options_result.Tx_Flow_out="All") {
		xon_choice:= "on"
	}
	else {
		xon_choice := "off"
	}
	if (Options_result.Rx_Flow_out = "DSR") {
		iDSR:="on"
	}
	else  iDSR:="off"
	;Rx_Flow= DTR or DTR_out = "DTR_Rx_Flow*" set DTR_bit ="hs"
	if ((Options_result.Rx_Flow_out = "DTR")  or (Options_result.DTR_out = "DTR_Rx_Flow*")) {
		DTR_bit :="hs"		  ;DTR pin has assumeed Ready_to_Receive Rx_flow control function, no DTR, no data allowed
	} ; * May not be supported by all drivers
	else{
		DTR_bit :=Options_result.DTR_out ;DTR_bit and RTS_bit on/off values are set directly by Gui
	}
	if(Options_result.RTS_out = "RTR_Flow*") {		;RTR_Flow* -RTS pin assumes Ready_to_Receive Rx_flow control function
		RTS_bit :="hs"
	}
	else if (Options_result.RTS_out = "Data_Rdy") {	;RTS is asserted when any data is ready to transmit, off otherwise.
		RTS_bit :="tg"
	}
	else {										;only On/Off remain
		RTS_bit:=Options_result.RTS_out
	}
	if (Options_result.Tx_Flow_out = "CTS") or (Options_result.Tx_Flow_out= "All")
	  oCTS:="on"
	else oCTS:="off"
	if (Options_result.Tx_Flow_out = "DSR") or (Options_result.Tx_Flow_out = "All")
	  odsr:="on"
	else oDSR:="off"

/*-----------------------------Configure a 'mode' command line from menu choices--------------------------------
*This program uses NEW format mode syntax
*	mode syntax -- com<m>[:] [baud=<b>] [parity=<p>] [data=<d>] [stop=<s>] [xon={on|off}]
*	[odsr={on|off}] [octs={on|off}] [dtr={on|off|hs}] [rts={on|off|hs|tg}] [idsr={on|off}]
*	[to={on|off}] is not used. Timeouts are set in RS232_Initialize()
*   RS232_Initialize() can recover from a bad comm parameter by returning to the com_config GUI until the user gets it right.
*/
RS232_Settings:=(Options_result.COM_Port_out
	":baud=" Options_result.Bit_Rate_out
	" data=" Options_result.N_bits_out
	" parity=" P_bit
	" stop=" Options_result.S_bits_out
	" dtr=" DTR_bit
	" rts=" RTS_bit
	" xon=" xon_choice
	" odsr=" oDSR
	" octs=" oCTS
	" idsr=" iDSR)

	;MsgBox("RS232_Settings " RS232_Settings,"Com_Port_Services. Line number " A_LineNumber)
	RS232_Port:=Options_result.COM_Port_out
	Com_cfg_Gui.Destroy()         			; we'll be back if config changesfail
	SetTimer(Modem_Leads,0)    				; Remove the 100mS timer to sample Modem leads
	if (  Old_Port!=RS232_Port) {    		; New port to open
		RS232_Close(RS232_FileHandle)   	; close the now open port
		if (X_Sync) {				; sync with external functions?
			Start_Program()             	; Restart Partner program to pick up new port
		}
	}
	;RS232_Initialize(RS232_Settings)       ; can create a circular reference!
	return									; new settings are not yet acted on.
	}

Com_cfg_GuiButtonCANCEL(*)	{	;Continue without Update to Configuration file
		ToolTip()            ; close any tooltip
		Com_cfg_Gui.Destroy()         ;done for now forgetaboutit
		SetTimer(Modem_Leads,0)    ; Remove the 100mS timer to sample Modem leads
		return                   ; end comconfig thread
	}

New_Rx_Flow(*)	{	;entr when the Rx Flow Control drop down list selection changes
	 Com_cfg_result:=Com_cfg_Gui.Submit(NoHide)
		if (Com_cfg_result.Rx_Flow_out ="DTR") {
			Msg_Var.SetFont "s10 cRed"
			Msg_Var.Value := "Caution ! Many drivers do not support this choice"
			DTR_Opt_DDL.Enabled := false
			RTS_Opt_DDL.Enabled := true
			RTS_Opt_DDL.Choose(RTS_Pick)
		}
		else if (Com_cfg_result.Rx_Flow_out = "RTR (RTS)")  {
			Msg_Var.SetFont "s10 cRed"
			Msg_Var.Value := "Caution ! Many drivers do not support this choice"
			RTS_Opt_DDL.Choose("RTR_Flow")
			RTS_Opt_DDL.Enabled := false
			DTR_Opt_DDL.Enabled := true
			DTR_Opt_DDL.Choose(DTR_Pick)
		}
		else {
			;MsgBox("DTR_PicK " DTR_Pick "RTS_Pick " RTS_Pick)
			Msg_Var.SetFont "s8 cblack"
			Msg_Var.Value:= ""
			DTR_Opt_DDL.Enabled := true
			DTR_Opt_DDL.Choose(DTR_Pick)
			RTS_Opt_DDL.Enabled := true
			RTS_Opt_DDL.Choose(RTS_Pick)
		}
		return
	}
New_Rts(*) {
		Com_cfg_result:=Com_cfg_Gui.Submit(NoHide)
		if (Com_cfg_result.RTS_out ="Data_Rdy") or (Com_cfg_result.RTS_out ="RTR_Flow*"){
			Msg_Var.SetFont "s10 cRed"
			Msg_Var.Value := "Caution ! Many drivers do not support this choice"
		  }
		else {
			Msg_Var.SetFont "s8 cblack"
			Msg_Var.Value := ""
			 }
		return
	}
;********** Timer landing that displays current modem output control leads in the comm port configuration GUI.
Modem_Leads()
	; Note that the target GUI gui is named, the process must designate the name in GuiControl parameter #1. i.e. Com_cfg_Gui:
	{
	  if (RS232_FileHandle >0)  {
		leads:=Get_CommModemLeads(RS232_FileHandle)     ; Read port input lead states
		; CTS_ON 0x0010  DSR_ON 0x0020  RING_ON 0x0040 RLSD_ON 0x0080
		if (leads & 0x0010 )
		  CTS_Lead_Disp.Opt "cgreen"
		else      CTS_Lead_Disp.Opt "cred"
		if (leads & 0x0020 )
		   DSR_Lead_Disp.Opt "cgreen"
		else       DSR_Lead_Disp.Opt "cred"
		if (leads & 0x0040 )
			 RI_Lead_Disp.Opt "cgreen"
		 else        RI_Lead_Disp.Opt "cred"
		if (leads & 0x0080 )
		   DCD_Lead_Disp.Opt "cgreen"
		else      DCD_Lead_Disp.Opt "cred"
	  }
	return
	}
}
;###################################### end com configure GUI ############################################################

;********Function to read the control leads of an open comm port  *****************
;   Return value is the sum of individual lead state values   **********************
;   CTS_ON 0x0010  DSR_ON 0x0020  RING_ON 0x0040 RLSD_ON 0x0080   *****************
;----------------------------------------------------------------------------------
Get_CommModemLeads(RS232_FileHandle) {
ModemStat:=Buffer(16)
Read_Result := DllCall("GetCommModemStatus"
	, "Uint",RS232_FileHandle
	, "Ptr",ModemStat
	,"Cdecl Uint")
Modem_Leads := NumGet(ModemStat,0, "UInt")		;reads an unsigned integer from the ModemStat buffer
												; CTS_ON 0x0010  DSR_ON 0x0020  RING_ON 0x0040 RLSD_ON 0x0080
  If (Read_Result = 0)
  {
     Error_detail:=GetErrorString(A_LastError)
    MsgBox "COM_Port Lead state read error `n Error Return value " A_LastError "`n Error message: " Error_detail "`nSerial Port lead States "  Modem_Leads , "Com Port Services Line number " A_LineNumber, "4096"
  }
Return Modem_Leads
}  ;********************************************************************************

InArr(haystack, needle) {			;returns index into haystack where needle if found, 0=not found.
	if !(IsObject(haystack)) or (haystack.Length = 0)
		return 0
	for index, value in haystack
		if (value = needle)
			return index
	return 0
}

;########################################################################
;###### Close RS23 COM Subroutine #######################################
;########################################################################
RS232_Close(FileHandle)  {

  CH_result := DllCall("CloseHandle", "UInt", FileHandle)
  If (CH_result != 1)       {       ; CloseHandle failed
    if (A_LastError !=6)    {       ; no message if "invalid file handle"
      Error_detail:=GetErrorString(A_LastError)
      MsgBox("Failed Dll CloseHandle `n Error code "  A_LastError "`n Error message: " Error_detail
	 ,"Com_Port_Services. Line number " A_LineNumber)
    }
  }
Global RS232_FileHandle:=0        ; invalidate FileHandle
  Return
}

RS232_Initialize(RS232_Mode)  {
  Global Config_Font
  Global Config_Message
  Global RS232_FileHandle
  ;
  ;###### Extract/Format the RS232 COM Port Number allowing port # greater then 9   ######

  retry_build_DCB:
  RS232_Temp:= StrSplit(RS232_Mode, ":" )    	;output should be a two element array,  RS232_Temp[1] =COMxy  and RS232_Temp[2] =(remainder of RS232_Mode)
  if (StrLen(RS232_Temp[1]) > 4)   {            ;So the valid names here are
    RS232_COM := "\\.\" RS232_Temp[1]           ; ... COM8  COM9   \\.\COM10  \\.\COM11  \\.\COM12 and so on...
  }
  else {
    RS232_COM := RS232_Temp[1]                  ; RS232_COM is the 'file open' name
  }
  ;MsgBox(" RS232_COM=" RS232_COM "`n DCB_Mode=" RS232_Temp[2] ,"Com_Port_Services. Line number " A_LineNumber)

  ;###### Build RS232 COM DCB ######
  ;Creates the structure that contains the RS232 COM Port number, baud rate,...

  Static DCB := Buffer(28, 0)					;Fill = 0x00
  NumPut "uint", 28,  DCB, 0					; put DCB size in 1st DWORD
  BCD_Result := DllCall("BuildCommDCB"
	  , "str", RS232_Temp[2]    				;RS232_Mode with COM # and : removed for BuildCommDCB.
	  , "Ptr", DCB)

  If (BCD_Result = 0)
  {
    msgResult:= MsgBox( "Failed Dll BuildCommDCB  BCD_Result=" A_LastError "`n Error message: "  GetErrorString(A_LastError) "`n COM parameters-" RS232_Mode
	   "`n`nClick RETRY to Check Com Config for unsupported options. CANCEL to exit" , "Com_Port_Services  Line Number " A_LineNumber , "R/C 4096")
    if (msgResult = "Retry")    {
      Config_Font:="s10 cRed"
      Config_Message:="Unsupported Comm Parameter"
      Com_Config_GUI(Com_inifile)              	; Allow user to fix parameters in (Global)  RS232_Mode using the GUI
      WinWaitClose("Com Port Configuration")   	; Wait for GUI to close. If error is not fixed we fail again
	  RS232_Mode:=RS232_Settings				; import new settings
	Goto retry_build_DCB						; see if user fixed fault
    }
    SAFE_EXIT()                                 ; cancel, give up but Comm port is not UP!
  }

  ;###### Open RS232 COM File if none open######
  If (RS232_FileHandle < 1)  { ; skip if a file is already open
		;Creates the RS232 COM Port File Handle,
		;see https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilea
		RS232_FileHandle := DllCall("CreateFile"
		,"Str", RS232_COM
		, "UInt", 0xC0000000		; Generic R/W access
		, "UInt", 3					; Shared R/W access (should be 0=exclusive access,?)
		, "UInt", 0					; Security attributes
		, "UInt", 3					; Opens a file or device, only if it exists.
		, "UInt", 0					; device attributes and flags, none assigned, (non-overlapped)
		, "UInt", 0
		, "Cdecl Int")
		if (RS232_FileHandle < 1)  	{					; Com Port failed to open
			Error_detail:=GetErrorString(A_LastError)
			if(A_LastError=2) {									; Can't find port
				WinWaitClose("Com Port Configuration")  		; Wait for GUI to close if open.
				Config_Font:="s10 cblue"
				Config_Message:="System can not find the requested Comm Port"
				Com_Config_GUI(Com_inifile)               		; Allow user to fix parameters in (Global)  RS232_Mode using the GUI
				WinWaitClose("Com Port Configuration") 			; Wait for GUI to close. If error is not fixed we fail again
				RS232_Mode:=RS232_Settings						; import new settings
				Goto retry_build_DCB							;see if user fixed fault
			}	;Effort to resolve this could result in on the fly user port# or DCB change...
			;Message box options, T5=5 sec. timer, 4096=always on top, R/C=Retry-Cancel box type
			msgResult:= MsgBox("Serial Port " RS232_COM " failed to Open. Error code " A_LastError "`nError message: " Error_detail
			"`nClick Cancel to Exit Serial Master.","Com_Port_Services. Line number " A_LineNumber,"T5 R/C 4096")
			if (msgResult = "Retry") {             ; This retry effort will keep trying the same settings...
				Goto retry_build_DCB             ; until the user clicks Cancel to exit.
			}
			if (msgResult = "Timeout") {
				RS232_Close(RS232_FileHandle)
				Goto retry_build_DCB
			}
		}
	}

  ;###### Set COM State ###### ;Sets the RS232 COM Port number, baud rate,...
	if (InStr(RS232_Temp[2], "xon=on")) {
		NumPut "UShort", 2048,DCB, 14			; Set XonLim.  Default RX buffer size is 4Kbytes.
		NumPut "UShort",512 ,DCB, 16			; Set XoffLim
		NumPut "uChar", 17,  DCB, 21			; Set Xon character
		NumPut "uChar", 19,  DCB, 22			; Set Xoff character
	}
  SCS_Result := DllCall("SetCommState"
	  , "UInt", RS232_FileHandle
	  , "Ptr", DCB)
  If (SCS_Result != 1)
  {
	Error_detail:=GetErrorString(A_LastError)
    MsgBox("Failed Dll SetCommState, SCS_Result=" A_LastError "`n Error message:" Error_detail "`nThe Script Will Now Exit."
	   ,"Com_Port_Services. Line number " A_LineNumber)
   SAFE_EXIT()
  }

  ;###### Create the SetCommTimeouts Structure ######
  ; ReadIntervalTimeout value of MAXDWORD, combined with zero values for both the ReadTotalTimeoutConstant
  ; and ReadTotalTimeoutMultiplier members, specifies that the read operation is to return immediately with the bytes
  ; that have already been received, even if no bytes have been received.  If WriteTotalTimeoutMultiplier and
  ; WriteTotalTimeoutConstant are both 0, Write never times out.
  ; Desired WriteTotalTimeoutMultiplier=(1Sec/baud_rate bits) x (10bits/Char) x (1000mS/Sec) = mSec/char or...
  ; (10,000 bits-mSec/char-Sec) x (1Sec./baud_rate bits.  or  10,000/baud_rate = mSec/char
  ; WriteTotalTimeoutMultiplier will be 90 for 110 baud and zero for baud rates above 9600.
  ; Transmit timeouts apply only to in-process write operations. A write operation may be Held Pending indefinitely
  ; by CTS or DSR output flow control.

  baud_start:=InStr(RS232_Mode, "baud=")+5
  baud_end:= InStr(RS232_Mode, A_Space, , baud_start)
  baud:=SubStr(RS232_Mode, (baud_start)<1 ? (baud_start)-1 : (baud_start), baud_end-baud_start)


  static ReadIntervalTimeout        := 0xffffffff               ;MAXDWORD...see above
  static ReadTotalTimeoutMultiplier := 0x00000000
  static ReadTotalTimeoutConstant   := 0x00000000
  static WriteTotalTimeoutMultiplier:= 0x00000000			; both write time out values = 0 then Tx timeouts not used
  Static WriteTotalTimeoutConstant  := 0x00000000        	; WriteTotalTimeoutConstant
  ;WriteTotalTimeoutMultiplier:=format("{:0#.8x}", 10000/baud) ; N byutes to write x WriteTotalTimeoutMultiplier (mSec) adds to
  ;WriteTotalTimeoutConstant:=format("{:0#.8x}", 5000)     ;convert 5 Sec. to long Hex
  ;MsgBox(" WriteTotalTimeoutConstant, " WriteTotalTimeoutConstant "`n WriteTotalTimeoutMultiplier " WriteTotalTimeoutMultiplier
  ; . "`n Line number " A_LineNumber)

  static Data := Buffer(20, 0) ; 5 * sizeof(DWORD) ;
  NumPut("UInt", ReadIntervalTimeout, Data, 0)
  NumPut("UInt", ReadTotalTimeoutMultiplier, Data, 4)
  NumPut("UInt", ReadTotalTimeoutConstant, Data, 8)
  NumPut("UInt", WriteTotalTimeoutMultiplier, Data, 12)
  NumPut("UInt", WriteTotalTimeoutConstant, Data, 16)

  ;### Set the RS232 COM Timeouts. COM port timeouts keep RS232_Read() from blocking the script if no input data is available ####

  SCT_result := DllCall("SetCommTimeouts"
	  , "UInt", RS232_FileHandle
	  , "Ptr", Data)
  If (SCT_result != 1)
  {
    MsgBox("Setting Serial Port Timeouts Failed. `nFailed Dll SetCommTimeouts, SCT_result=" A_LastError "`n Error message: "  GetErrorString(A_LastError)
	  . " `nThe Script Will Now Exit.","Com_Port_Services. Line number " A_LineNumber)
    SAFE_EXIT()
  }
;show_DCB(DCB, A_LineNumber)
  ;MsgBox( RS232_COM " should be open for business`n RS232_FileHandle " RS232_FileHandle "`n RS232_Mode=" RS232_Mode
  ;. ,"Com_Port_Services. Line number " A_LineNumber))
  return RS232_FileHandle
}  ;##### Com Port is open for business##########################################

;####################################################################################
;######           Write to RS232 COM Subroutines ####################################
;### Data_Type = 1 (default) or H or h indicates a comma delimited string of ASCII ##
;###   character codes, The individual codes may be decimalnumbers 0-255 or in  #####
;##  ASCII-Hex form, 0x00-0xFF. A minimum of 1 code is required.				#####
;### All other Data_Type values process the input string as a 'normal' string   #####
;####################################################################################
RS232_Write(RS232_FileHandle,Tx_Message,Data_Type:=1)  {
  Global Config_Font
  Global Config_Message
  Bytes_Sent:=0
  if(Data_Type !=1) and  (SubStr(Data_Type, 1, 1) !="H")    {
		Data_Length := StrLen(Tx_Message) 	; This is a 'Normal" format string
		Data := Buffer(Data_Length, 0xFF) ;Set the Data buffer size, prefill with 0xFF.
		Loop Data_Length
		{
		  Msg_Byte:=Ord(SubStr(Tx_Message, A_Index, 1)) ; pick 1 char from string and convert to ASCII
		  NumPut("UChar", Msg_Byte, Data, (A_Index-1)) ;Write the Message byte into the Data buffer
		}
	}
  else {                                    ;Data_Type =1 or H/h  Process as an ASCII string
    StrReplace(Tx_Message, ",", ",", , &Data_Length)   ;Count the # of commas, don't remove them
    Data_Length:=Data_Length+1                  ; Data elements is (# of commas)+1
    Data := Buffer(Data_Length, 0xFF)     		; Set the Data buffer size, prefill with 0xFF.
    Loop Parse, Tx_Message, ","                 ;
    {
       NumPut("Char", A_LoopField, Data, (A_Index-1))    ;Write the Message into the Data buffer
    }
  }

  ;###### Write the data to the RS232 COM Port ######
  WF_Result := DllCall("WriteFile"
	  , "UInt", RS232_FileHandle
	  , "Ptr", Data					;pointer to the buffer containing the data to be written
	  , "UInt", Data_Length			; Bytes to send
	  , "UInt*", &Bytes_Sent
	  , "Int", 0)
  If (WF_Result = 0 or Bytes_Sent != Data_Length) {
    Error_detail:=GetErrorString(A_LastError)
    if ((Bytes_Sent != Data_Length) and (A_LastError =0)) {
      Config_Font:="cRed"                           ; Keep default size
      Config_Message:="  Serial Port Transmit Failed Check TX_Flow Control"
      Com_Config_GUI(Com_inifile)                  	; Allow user to fix parameters in (Global)  RS232_Settings using the GUI
      WinWaitClose("Com Port Configuration")      	; Wait for GUI to close. If error is not fixed we fail again
	  RS232_Initialize(RS232_Settings)       		; Apply new settings
      ; Need to execute  ClearCommError function and look into the COMSTAT structure
      ; See https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-clearcommerror
	  ;MsgBox("Back from Com Port Configuration")
    }
    else {
		MsgBox("Failed Dll WriteFile to RS232 COM, result=" A_LastError "`n Error Message: " Error_detail
		. "`nFileHandle-" RS232_FileHandle "`nData Length=" Data_Length "`nBytes_Sent=" Bytes_Sent , "Com_Port_Services. Line number " A_LineNumber)
	}
  }
  return Bytes_Sent
}  ;###########  End Write to RS232 COM Subroutines######################################################

;#########################################################################################################
;###### Read from RS232 COM Subroutines  #################################################################
;## A Hex-ASCII string of receiced characters is returned and RS232_Bytes_Received is updated    #########
;#########################################################################################################

RS232_Read(XRS232_FileHandle, Num_Bytes, &RS232_Bytes_Received) {
Global RS232_FileHandle						; Opened for Non-Overlapped I/O
Global RS232_Port
  retry_count:=0
  Retry_read:
  ;Bytes_Received:=0 						;Part of the work-arround to prevent crash and burn on ExitApp
  Data := Buffer(Num_Bytes, 0x55)  			;Set the requested Data buffer size, prefill with 0x55 = ASCII character "U"
  ;###### Read the data from the RS232 COM Port ######
  Read_Result := DllCall("ReadFile"
  , "UInt", RS232_FileHandle
  , "Ptr", Data								; Pointer to the buffer that receives the data
  , "Int", Num_Bytes						; Maximum number of bytes to read
  ;, "Uint*", &Bytes_Received				; Pointer to the number of bytes read
  , "Int*", &RS232_Bytes_Received			; Original Pointer to the variable that receives the number of bytes read
  , "Int", 0)								; This parameter is optional and applies only if open for Overlapped I/O.
  ;RS232_Bytes_Received:=Bytes_Received		;Part of the work-arround to prevent crash and burn on ExitApp
  ;If the function succeeds, the return value is nonzero (TRUE).
  ;If the function fails, or is completing asynchronously, the return value is zero (FALSE).
  ; Asynchronous completion is not expected.

  ;MsgBox("RS232_FileHandle=" RS232_FileHandle "`nRead_Result=" Read_Result "`nBytes_Received=" RS232_Bytes_Received
  ;. "`nData="  Format("{:X}", NumGet(Data, 0 , "UChar")) "`n Line number " A_LineNumber)
  If (Read_Result = 0)
  {
	;MsgBox("Serial port read Error, Serial Port " RS232_Port "`n Error Return value " A_LastError "`n Error message: " Error_detail "`n Line number " A_LineNumber)
    if (retry_count < 10)  {        ; read recuvery after computer wakes from sleep
      Sleep(250)                     ; delay needed?
      retry_count:= retry_count +1  ; This will not work for most other reasons that a read will fail
      Goto Retry_read
    }
    Error_detail:=GetErrorString(A_LastError)
    ;Message box options, 4096=always on top, 5 is Retry-Cancel box type
    msgResult := MsgBox "Serial port read Error, Serial Port " RS232_Port  "failed with 10 retrys.`n Error Return value " A_LastError "`n Error message: " Error_detail " `nClick Cancel to Exit Serial Master.`nClick Retry to re-open Comm Port.","Com Port Services Line number " A_LineNumber ,"4096 R/C  t5"
    if (msgResult = "Retry")
      Goto Re_Open_Port
    if (msgResult = "Timeout")
      Goto Re_Open_Port
    SAFE_EXIT()
    Re_Open_Port:
    {
    RS232_Close(RS232_FileHandle)         	;close the now open port
		RS232_Initialize(RS232_Settings)     	; re-open port, sets a new global RS232_FileHandle
		;MsgBox("Retry settings " RS232_Settings "`nRETRY RS232_FileHandle " RS232_FileHandle "`n Line number " A_LineNumber)
		retry_count:=0
		Goto Retry_read
    }
  } ;end read error process
  if (retry_count >0 )  {
    MsgBox("Retry count " retry_count  "Serial Port " RS232_Port " Read restored" , "Com Port Services line number  " A_LineNumber , "4096 T6")
  }
  ;###### Format the received data ######
  ;This loop is necessary to successfully pass a received NULL (0x00) character from the serial interface
  ; using a string. The 0x00 character is used as the variable length string terminator and as such
  ; any binary zero stored in a variable by a function will hide all the data following the zero.  That is, such data
  ; cannot be accessed by string functions.
  ; Note that the Dll passes the data as a buffer with length and buffer start address.
  ; The loop converts each byte of the read data buffer to a 2 character Hex-ASCII code that is
  ; appended to Data_HEX for each byte received from the serial port.  e.g. Data = 00 => Data_HEX =3030 or
  ; Data = 1F => Data_HEX =3146.  With this encoding of received data, any byte value received can be passed
  ; in the encoded string.
  Data_HEX :=""
 Loop RS232_Bytes_Received
	{    ;First byte in the Rx FIFO ends up at position 0
		Data_HEX_Temp := Format("{:X}", NumGet(Data, A_Index-1 , "UChar")) ;Convert to HEX notation byte-by-byte
		If (StrLen(Data_HEX_Temp) =1) {         ;If there is only 1 character then add the leading "0'
		  Data_HEX_Temp := "0" Data_HEX_Temp
		}
		Data_HEX.=Data_HEX_Temp    ;Put it all together
	}
  Return Data_HEX
}
Show_DCB(temp_DCB, called_from)	{				; Opens a messagebox and displays DCB contents
/* Variable values preceeded by 'f' are bits within a double word. fBinary is the LSB
 * 0* DWORD DCBlength;
 *04* DWORD BaudRate			:   ; allowed values -- 110, 300 x 2^n where n =0-7, 14400, 57600,115200, 128000 or 256000
 *08* DWORD fBinary				: 1 ; TRUE, binary mode is required.( but not reported...)
 *	  DWORD fParity				: 1 ; TRUE, parity checking is performed and errors are reported.
 *	  DWORD fOutxCtsFlow		: 1 ; TRUE, if CTS is turned Off, output is suspended until CTS is turned On.
 *	  DWORD fOutxDsrFlow		: 1 ; TRUE, if DSR is turned off, output is suspended until DSR is turned On.
 *	  DWORD fDtrControl			: 2 ; 00 = DTR Off  01 = DTR On  02 = DTR handshaking
 *	  DWORD fDsrSensitivity		: 1 ; TRUE, driver ignores any bytes received, unless the DSR modem input line is high.
 *	  DWORD fTXContinueOnXoff	: 1 ; TRUE, transmission continues after the input buffer has come within XoffLim
 *09  DWORD fOutX				: 1 ; TRUE, transmission stops when XoffChar is received and re-starts on XonChar is received.
 *	  DWORD fInX				: 1 ; TRUE, XoffChar sent when input buffer > XoffLim and XonChar sent when input buffer <XonLim.
 *	  DWORD fErrorChar			: 1 ; TRUE and fParity TRUE, received characters with parity errors are replaced by ErrorChar.
 *	  DWORD fNull				: 1 ; TRUE, null bytes are discarded when received.
 *	  DWORD fRtsControl			: 2 ; 00 = RTS Off  01 = RTS On  02 = RTS handshaking 03 = RTS on if bytes are available to send.
 *	  DWORD fAbortOnError		: 1 ; TRUE, the driver terminates all operations with an error status if an error occurs.
 *	  DWORD fDummy				: 17; Reserved; do not use.
 *12* WORD  wReserved     		:   ; Reserved; must be zero.
 *14* WORD  XonLim				:   ; The minimum number of bytes in the input buffer before XonChar is sent.
 *16* WORD  XoffLim				:   ; The minimum number of free bytes in the input buffer before XoffChar is sent.
 *18* BYTE  ByteSize			:   ; The number of bits in the byte
 *19* BYTE  Parity				:   ; 0 =NOPARITY  1 =ODDPARITY 2 = EVENPARITY 3 =MARKPARITY 4=SPACEPARITY
 *20* BYTE  StopBits			:   ; The number of stop bits to be used.  0 =1 stop bit. 1=1.5 stop bit. 2 =2 stop bits.
 *21* char  XonChar				:   ; The value of the XON character for both transmission and reception.
 *22* char  XoffChar			:   ; The value of the XOFF character for both transmission and reception.
 *23* char  ErrorChar			:   ; The value of the character used to replace bytes received with a parity error.
 *24* char  EofChar				:   ; The value of the character used to signal the end of data.
 *25* char  EvtChar				:	; The value of the character used to signal an event.
 *27| WORD  wReserved1			:   ; Reserved; do not use.
 */
GCS_Result := DllCall("GetCommState"
  , "UInt", RS232_FileHandle
  , "Ptr", temp_DCB)
If (GCS_Result != 1)
{
Error_detail:=GetErrorString(A_LastError)
MsgBox("Failed Dll GetCommState, GCS_Result=" A_LastError "`n Error message:" Error_detail "`nThe Script Will Now Exit." , "Com Port Services Error. Line number " A_LineNumber,)
}
 DCB_Bitfield1:=NumGet(temp_DCB, 8,"UChar")
 DCB_Bitfield2:=NumGet(temp_DCB, 9,"UChar")
MsgBox ("DCB length " NumGet(temp_DCB, 0,"Uint") "`nBaud rate`t" NumGet(temp_DCB, 4,"Uint") "`nParity checking`t" DCB_Bitfield1 >> 1 & 0x01
 . "`nOutxCtsFlow`t" DCB_Bitfield1 >> 2 & 0x01 "`nOutxDsrFlow`t" DCB_Bitfield1 >> 3 & 0x01 "`nDtrControl`t" DCB_Bitfield1 >> 4 & 0x03
 . "`nDsrSensitivity`t" DCB_Bitfield1 >> 6 & 0x01 "`nContinueOnXoff`t" DCB_Bitfield1 >> 7 & 0x01 "`nOutX`t`t" DCB_Bitfield2 & 0x01
 . "`nInX`t`t" DCB_Bitfield2 >> 1 & 0x01 "`nErrorChar`t" DCB_Bitfield2 >> 2 & 0x01 "`nNull`t`t" DCB_Bitfield2 >> 3 & 0x01
 . "`nRtsControl`t" DCB_Bitfield2 >> 4 & 0x03 "`nAbortOnError`t" DCB_Bitfield2 >> 6 & 0x01 "`nXonLim`t`t" NumGet(temp_DCB, 14,"Ushort")
 . "`nXoffLim`t`t" NumGet(temp_DCB, 16,"Ushort") "`nByteSize`t`t" NumGet(temp_DCB, 18,"UChar") "`nParity`t`t" NumGet(temp_DCB, 19,"UChar")
 . "`nStopBits`t`t" NumGet(temp_DCB, 20,"UChar") "`nXonChar`t`t" NumGet(temp_DCB, 21,"UChar") "`nXoffChar`t`t" NumGet(temp_DCB, 22,"UChar")
 . "`nErrorChar`t" NumGet(temp_DCB, 23,"UChar") "`nEofChar`t`t" NumGet(temp_DCB, 24,"UChar")
 . "`nEvtChar`t`t" NumGet(temp_DCB, 25,"UChar") ) , "Com_Port_Services  Line number " called_from

}
