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
;msgbox("This program  is: AutoHotKey version " A_AhkVersion "`n line number " A_LineNumber)
;################### Serial_Master_Menu_V2.ahk ###################
; Last_edit:="HVP 09/30/2023 ; Gui title was "Serial Master"
; 09/19/2023 ; MenuHandler(init_only:=false,*) sets Global vatiables only when init_only=true or "OK" exit
; 09/17/2023 replaced WinActivate(Partner_ID, , Start_Delay) with WinWaitActive(Partner_ID, , Start_Delay) in Launch code.
;   Start_Delay is now global
; 08/20/2023"   ;Start_Program() and partner_app="None" no longer jumps to File_Start_Done: &
; Part_Change(*) was causing Partner_App to change before OK button. New_Partner_App was Partner_App.
; 07/05/2023"      ;Filename was Master_Menu_V2.ahk
; Version_Number:= "V2.0"
; This program is intended to be an #include for the Serial_MasterV2
; The Program Options Gui menu display allows the user to select an application program, Partner Application, and
; various Serial_Master parameters. The parameter selections for each Partner Application are saved in Serial_Master.cfg
; as a profile.  When the Gui is displayed the saved profile choices are displayed, and if unchanged are used to configure
; Serial Master operation.  Each Partner Application has a saved profile and once set up will reduce the number of clicks
; needed to configure Serial Master for a Partner Application.
; The functions included provide the following services:
; MenuHandler(init_only:=false,*)	Provides global variable initialization and/or a Serial Master Program Options Gui menu display
;   init_only=true: "Initialize only"  Reads Serial_Master.cfg and returns Global variables.
;   init_only=false: Opens 'Serial Master' Gui for editing AND returns INITIAL Global variables.
;   On Gui OK or Close:         Erases Gui, Saves edit results in Serial_Master.cfg, updates Global variables and Calls Start_Program().
;   On Gui CANCEL or Escape:	Erases Gui and returns with Global variables and Serial_Master.cfg unchanged.
;
; Start_Program()		The application profile is read from Serial_Master.cfg	using Partner_App as the profile key.
;						If no Companion file is identified the System is examined for Open copies of the Partner Application.
;							If found the Partner Application is Re-activated, if not found the Partner Application is started.
;						If a Companion file is identified in the profile and does not exist, the file is created.
;						If a Companion file is identified, the System is examined for open copies of the Companion file:
;							None found: Start Partner Application and Open file.
;							Companion file found is open in Partner Application, Re-activate Partner Application. Else,
;							Companion file found, User must Choose:	a) Start Partner Application and Open file
;																	b) Re-activate application that has file open
;																	c) Return to Program Options Gui menu
;
;

MenuHandler(init_only:=false,*) {
;
static NoHide:=0
; Global variables marked R/O in comment are 'Read Only', all others are set from Serial_Master.cfg
; Global variables marked * in comment are set in init_only mode and exit via OK button
;global RS232_Port              ; R/O Port number (COM'n') set by Com Port Gui
Global Partner_App              ; * Partner Program name
Global GIDEI_Decoder            ; * GIDEI_Decoder selection, see 'BI_Decoder_opt_list'
Global Key_List                 ; * Key_List is used by the GIDEI_Decoder this adds/removes "Volume_Mute.Volume_Down.Volume_Up."
Global Key_Mode                 ; * Keyboard data mode selection see 'BI_Keybd_opt_list'
;Global App_exe                 ; full path to partner App executable file. May be absent for App in the default Path
Global config_font              ; Font used in Gui Message area
Global Config_Message           ; Message text displayed in Gui Message area. IF "Initialize", only "Global' variables are set

; The following variables are Biult-in and may be used as 'Last Resort' option for failed Serial_Master.cfg file.
Global BI_Keybd_opt_list:="Off|Serial|Both|Force"
Global BI_Start_opt_list:="Launch|Launch, Open File|Wait, User start"
Global BI_File_name_list:="Console|DataLog|TextFile|None"
Global BI_File_ext_list:=".txt|.rtf|.doc|.docx|.csv|.xlsx|None"
Global BI_Start_Delay_List:="1|2|5|10|20|50|100"
Global BI_App_profile:=",3,2,1,1,1,1,2"
static BI_Decoder_opt_list:="On+Vol|On|Off"
static BI_Partner_Choice:= 1   ;Built-in default to Notepad
static BI_App_list:="Notepad|WordPad|Word|Excel|Media Center|None"
; First read config file to get the number of the partner program
Partner_Num := IniRead("Serial_Master.cfg", "Program_opt", "Partner_prog","ERROR")
  if ((Partner_Num= "ERROR") or(Partner_Num= "" )) {
    Partner_Num := IniRead("Serial_Master.cfg", "Default", "Partner_prog","ERROR") ;if ERROR, the config file should be presumed missing
    error_msg:= "Default Partner Number"
    if ((Partner_Num= "ERROR") or (Partner_Num= "" )) {
      Partner_Num:=BI_Partner_Choice
          error_msg:= "Built-in Partner Number"
    }
    Config_Default_Msg("Serial_Master.cfg", "Partner program", error_msg Partner_Num, A_LineNumber)     ; Pop message
  }  ;Next read config file to get the App_list
App_list := IniRead("Serial_Master.cfg", "Program_opt", "Prog_list","ERROR")    ;Read the program list
if ((App_list= "ERROR") or (App_list= "" )) {                  ; Pop message, use built-in
  Config_Default_Msg("Serial_Master.cfg", "Program list", "Built in list", A_LineNumber)
  App_list:=BI_App_list                     ; Serial_Master.cfg does not have a default list
}
App_array:=StrSplit(App_list,"|")           ; Split the App_list into an array of app names
App_profile := IniRead("Serial_Master.cfg", "Profile", App_array[Partner_Num],"ERROR")    ;read the App_profile useing the Partner_App
  if ( (App_profile= "ERROR") or ( App_profile= "" )) {   ; Pop message, use default
    App_profile := IniRead("Serial_Master.cfg", "Default", App_array[Partner_Num],"ERROR")   ;if this is ERROR, use default profile
    error_msg:= "Default profile"
    if ((App_profile= "ERROR") or (App_profile= "" )) {
      App_profile:=BI_App_profile           ; use last resort profile
      error_msg:= "Built-in profile"
    }
    Config_Default_Msg("Serial_Master.cfg", "profile", error_msg, A_LineNumber)
  }
profile_array:=StrSplit(App_profile, ",")   ; Split the App_profile into an array of initial Config choices
App_exe:=profile_array[1]
Keybd_opt_array:=StrSplit(BI_Keybd_opt_list,"|") ; Split the Keybd_opt into an array of Keybd_opt
Start_opt_array:=StrSplit(BI_Start_opt_list,"|")
Decoder_opt_array:=StrSplit(BI_Decoder_opt_list,"|")
if (init_only=true) {
  Partner_App:=App_array[Partner_Num]         ; now set the name of this Partner program
  ; profile_array paramater order: 1:App_exe, 2:Keyboard data, 3:Start Option 4:port# in file name 5:FileName
  ;                                6: File_ext 7: decoder option 8:Start_Delay
  Key_Mode :=Keybd_opt_array[profile_array[2]]
  GIDEI_Decoder:=Decoder_opt_array[profile_array[7]]
  if (profile_array[7] !=1)                   ; Not On+Vol, remove the volume control keys if they exist
    {
      Key_List:= StrReplace(Key_List,"Volume_Mute.Volume_Down.Volume_Up.")
    }
  else if !(InStr(Key_List, "Volume_Mute"))   ; if not already there, add them
    {
      Key_List.="Volume_Mute.Volume_Down.Volume_Up."                ; Append volume control keys
    }
   ; MsgBox("Key_List " Key_List "`nLine Number " A_LineNumber )
  ; MsgBox( "Initialize only `nPartner_App= " Partner_App "`n Key_Mode= " Key_Mode "`n GIDEI_Decoder= " GIDEI_Decoder "`n Line number " A_LineNumber)
  return
}

; now read the File_name_list
File_name_opt := IniRead("Serial_Master.cfg", "Program_opt", "File_Name_List","ERROR")
  if ((File_name_opt= "ERROR") or (File_name_opt= "" )) {   ; Pop message, use built-in
    Config_Default_Msg("Serial_Master.cfg", "File_Name_List", "Built in list", A_LineNumber)
    File_name_opt:=BI_File_name_list                 ; Serial_Master.cfg does not have a default list
  }
File_name_array:=StrSplit(File_name_opt,"|")
File_ext_opt := IniRead("Serial_Master.cfg", "Program_opt", "File_ext_List","ERROR")
  if ((File_ext_opt= "ERROR") or (File_ext_opt= "" )) {   ; Pop message, use built-in
    Config_Default_Msg("Serial_Master.cfg", "File_ext_List", "Built in list", A_LineNumber)
    File_ext_opt:=BI_File_ext_list                 ; Serial_Master.cfg does not have a default list
  }
File_ext_array:=StrSplit(File_ext_opt,"|")

Start_time_opt := IniRead("Serial_Master.cfg", "Program_opt", "Start_Time_List","ERROR")
  if ((Start_time_opt= "ERROR") or (Start_time_opt= "" )) {   ; Pop message, use built-in
    Config_Default_Msg("Serial_Master.cfg", "Start_Time_List", "Built in list", A_LineNumber)
    Start_time_opt:=BI_Start_Delay_List                 ; Serial_Master.cfg does not have a default list
  }
 Start_Delay_array:=StrSplit(Start_time_opt,"|")


Options_width:=450
OK_X_pos:=Options_width-130
Options_height :=300 ;Sets the vertical size of the box
Exit_Button_x :=Options_width - 60
Exit_Button_y :=Options_height -120

Message_width:=Options_width-32
Program_options := Gui(,"Serial Master Options")
Program_options.Opt("-SysMenu")   		; removes icon and control buttons from title bar
Program_options.Opt("+alwaysontop") 	;another window will not hide this Gui
Program_options.MarginX:= 10            ;replaces Program_options.Margin ("10", "10")form V1
Program_options.MarginY:= 10
Program_options.SetFont "s14 norm"
Program_options.Add("Text", "xm y0", "Serial Master Program Options")
Program_options.SetFont "s12 norm"
Program_options.Add("Button", "x" OK_X_pos " y0 default", "OK").OnEvent("Click", Program_optionsButtonOK) ; call when "OK" is clicked
Program_options.Add("Button","x+2 y0", "CANCEL").OnEvent("Click", Program_optionsButtonCANCEL)
Program_options.Add("Button", "x" Exit_Button_x " y" Exit_Button_y, "EXIT").OnEvent("Click", Program_optionsExit)
Program_options.SetFont "s10"
Program_options.Add("GroupBox", "x6 y30 w130 h50", "Partner Program")
Program_options.Add("DropDownList", "x15 y50 w110 altsubmit vPart_Prog  Choose" Partner_Num, App_array).OnEvent("Change", Part_Change)
;altsubmit returns the choice list position number in Part_Prog instead of entry text
Program_options.Add("GroupBox", "x16 y85 w100 h50", "Keyboard Input")
Keybd_opt_DDL:=Program_options.Add("DropDownList", "x25 y105 w60 altsubmit vK_Data Choose" profile_array[2], Keybd_opt_array)
Program_options.Add("GroupBox", "x140 y30 w140 h50", "Start Option")
Start_opt_DDL:=Program_options.Add("DropDownList", "x146 y50 w130 altsubmit vStart_opt Choose" profile_array[3], Start_opt_array)
Program_options.Add("GroupBox", "x140 y85 w140 h70", "File Name")
Include_Port_Box:=Program_options.Add("CheckBox", "x146 y105 vport_n Checked" profile_array[4], "Include Com Port #" )
File_name_opt_DDL:=Program_options.Add("DropDownList", "x146 y125 w75 altsubmit vFile_name_opt Choose" profile_array[5], File_name_array)
File_ext_opt_DDL:=Program_options.Add("DropDownList", "x220 y125 w55 altsubmit vFile_ext_opt Choose" profile_array[6], File_ext_array)
Program_options.Add("GroupBox", "x16 y155 w110 h50", "GIDEI Decoder")
Decoder_opt_DDL:=Program_options.Add("DropDownList", "x30 y175 w80 altsubmit vDecoder_opt Choose" profile_array[7], Decoder_opt_array)
Program_options.Add("GroupBox", "x140 y155 w110 h50", "Start-up Delay")
Start_timer_opt_DDL:=Program_options.Add("DropDownList", "x146 y175 w80 altsubmit vStart_timer_opt Choose" profile_array[8], Start_Delay_array)
Program_options.Add("GroupBox", "x16 y205 w" . Message_width . " h50", "Message Area")
Program_options.SetFont config_font
Program_options.Add("text", "vConfig_Msg_Var xm y225 w" Message_width " h70", Config_Message)
Program_options.SetFont("")             ;restore system default font
Program_options.Show("w" Options_width " h" Options_height)             ;V2 change, Gui Nane is in GUI(,"Serial Master")
Program_options.OnEvent("Close",Program_optionsButtonOK)
Program_options.OnEvent("Escape",Program_optionsButtonCANCEL)
Config_Message:=""
return



Part_Change(*)  {            ;user changed the partner program selection, Apply a new profile to variables
  Options_result := Program_options.Submit(NoHide)          ;NoHide=0
  New_Partner_App:=App_array[Options_result.Part_Prog]          ; now find the name new Partner program
  App_profile := IniRead("Serial_Master.cfg", "Profile", New_Partner_App,"ERROR") ;read the App_profile useng the New_Partner_App
  if ( (App_profile= "ERROR") or ( App_profile= "" )) {                         ; Pop message, use default
    App_profile := IniRead("Serial_Master.cfg", "Default", "Profile_Default")   ;if this is ERROR, use default profile
    Config_Default_Msg("Serial_Master.cfg", "profile", App_profile, A_LineNumber)
  }
  profile_array:=StrSplit(App_profile, ",")  ; Split the App_profile into an array of Config choices
  ;MsgBox("partner program change " New_Partner_App "`n  New Profile " App_profile "`n Line number " A_LineNumber)
  App_exe:=profile_array[1]
  ;in the following '0+' forces the array value to look like a Number....
  Keybd_opt_DDL.Choose(0+profile_array[2])                  ; Apply the new profile Gui
  Start_opt_DDL.Choose(0+profile_array[3])
  Include_Port_Box.Value:=0+profile_array[4]
  File_name_opt_DDL.Choose(0+profile_array[5])
  File_ext_opt_DDL.Choose(0+profile_array[6])
  Decoder_opt_DDL.Choose(0+profile_array[7])
  Start_timer_opt_DDL.Choose(0+profile_array[8])
  return
  }

Program_optionsExit(*)	{	                                ; Exit program button clicked
  Program_options.Destroy()
  SAFE_EXIT()
  exit
}
Program_optionsButtonCANCEL(*)	{	                        ; Abort or EscapeGUI without Update
  ;MsgBox("Cancel ButtonPressed")
  Program_options.Destroy()                                 ;done for now forgetaboutit
  return
}

Program_optionsButtonOK(*)   {                              ;Save the new program configuration...Partner_prog & Profile
    Options_result := Program_options.Submit(NoHide)         ;NoHide=0   Get new paramaters ...
    Partner_App:=App_array[Options_result.Part_Prog]        ; name of profile key (App name)
    new_prof:=App_exe "," Options_result.K_Data "," Options_result.Start_opt "," Options_result.port_n "," Options_result.File_name_opt "," Options_result.File_ext_opt "," Options_result.Decoder_opt "," Options_result.Start_timer_opt
    Key_Mode:=Keybd_opt_array[Options_result.K_Data]

    GIDEI_Decoder:=Decoder_opt_array[Options_result.Decoder_opt]
    if (Options_result.Decoder_opt!=1)                      ; Not On+Vol, remove the volume control keys if they exist
      {
        Key_List:= StrReplace(Key_List,"Volume_Mute.Volume_Down.Volume_Up.")
      }
    else if !(InStr(Key_List, "Volume_Mute"))               ; if not already there, add them
      {
        Key_List.="Volume_Mute.Volume_Down.Volume_Up."      ; Append volume control keys
      }
    ; MsgBox( "On OK `nPartner_App= " Partner_App "`n Key_Mode= " Key_Mode "`n GIDEI_Decoder= " GIDEI_Decoder "`n Line number " A_LineNumber)
    try {    ;V2 An OSError is thrown on failure
      IniWrite(new_prof, "Serial_Master.cfg", "Profile", Partner_App)                         ;Save a profile line in Serial_Master.cfg
      IniWrite(Options_result.Part_Prog, "Serial_Master.cfg", "Program_opt", "Partner_prog")  ;Save active program in Serial_Master.cfg
    }
    catch  {
      MsgBox("Serial_Master.cfg write failed Line number " A_LineNumber)
    }
    Program_options.Destroy()                               ;done for now, forgetaboutit
    ;*** Program Options GUI closed and settings are saved in Serial_Master.cfg
    ; Next, start the new configuration
    Start_Program()
    return
  } ; End Program_optionsButtonOK()
}   ; End of Program_options Gui processes

Start_Program() {

    global  Partner_App         ; Must be valid on entry!
    global  Partner_ID
    global  Partner_PID
    Global Start_Delay              ; Application max start-up delay configuration parameter
    ; Called as part of the cold start process or continuation of Program Options Gui close

    App_profile := IniRead("Serial_Master.cfg", "Profile", Partner_App,"ERROR")    ;read the App_profile useing the Partner_App
      if ( (App_profile= "ERROR") or ( App_profile= "" )) {             ; Pop message, use default
        App_profile := IniRead("Serial_Master.cfg", "Default", Partner_App,"ERROR")   ;if this is ERROR, use default profile
        error_msg:= "Default profile"
        if ((App_profile= "ERROR") or (App_profile= "" )) {
          App_profile:=BI_App_profile       ; use last resort profile
          error_msg:= "Built-in profile"
        }
        Config_Default_Msg("Serial_Master.cfg", "profile", error_msg, A_LineNumber)
      }
    profile_array:=StrSplit(App_profile, ",")  ; Split the App_profile into an array of initial Config choices
    App_exe:=profile_array[1]
    ; profile_array paramater order: 1:App_exe, 2:Keyboard data, 3:Start Option 4:port# in file name 5:FileName
    ;                                6: File_ext 7: decoder option 8:Start_Delay
    if (App_exe ="")   {
      App_exe :=Partner_app ".exe"                                   ; if App_exe is not specified use Partner_app
    }
    App_exe_short:=App_exe
    Last_slash:=InStr(App_exe, "\", , -2)                             ; find end of Path info, if any
    if Last_slash  {
      App_exe_short:=SubStr(App_exe, Last_slash+1)  ; keep only the File name
    }
    App_exe_short:=SubStr(App_exe_short, 1, InStr(App_exe_short, ".", , 1)+3) ; trim anything that follows '.exe'
    ;MsgBox("App_exe_short " App_exe_short)
    Start_Array:=StrSplit(BI_Start_opt_list,"|")                    ; Start Options= "Launch|Launch, Open File|Wait, User start"
    Start_opt:=Start_Array[profile_array[3]]                        ; Start_opt now has a name
    File_name:=""
    if(profile_array[4])        {                                   ; include port# in File name? 0= no, 1= yes
      File_name:=RS232_Port "_"
    }
    File_name_list := IniRead("Serial_Master.cfg", "Program_opt", "File_name_list","ERROR") ;read the File_name_list
    if ( (File_name_list= "ERROR") or ( File_name_list= "" )) {    ; Pop message, use default
      Config_Default_Msg("Serial_Master.cfg","profile", File_name_list,A_LineNumber)
      File_name_list := IniRead("Serial_Master.cfg", "Default", "File_name_list","ERROR")  ;if this is ERROR, usebuilt-in File_name_list
      if ((File_name_list= "ERROR") or (File_name_list= "" ))
        File_name_list:=BI_File_name_list                          ;last resort File_name_list
    }
    File_name_array:=StrSplit(File_name_list, "|")                  ; Split the File_name_list into an array
    File_name.=File_name_array[profile_array[5]]                    ; Append profile choice file name
    File_ext_list := IniRead("Serial_Master.cfg", "Program_opt", "File_ext_list","ERROR") ;read the File_extension list
    if ((File_ext_list= "ERROR") or (File_ext_list= "" )) {        ; Pop message, use default
      Config_Default_Msg("Serial_Master.cfg","profile", File_ext_list, A_LineNumber)
      File_ext := IniRead("Serial_Master.cfg", "Default", "File_ext_list","ERROR")  ;if this is ERROR, usebuilt-in File_name
      if ((File_ext_list= "ERROR") or (File_ext_list= "" ))
        File_ext_List:=BI_File_ext_list                                  ;last resort File_ext
    }
    File_ext_array:=StrSplit(File_ext_list, "|")                    ; Split the File_extension_list into an array
    File_name.=File_ext_array[profile_array[6]]                     ; Append profile choice file ext list
    Start_Delay_List := IniRead("Serial_Master.cfg", "Program_opt", "Start_Time_List","ERROR") ;read the file start timer list
    if ((Start_Delay_List= "ERROR") or (Start_Delay_List= "" )) {        ; Pop message, use default
      Config_Default_Msg("Serial_Master.cfg","profile", Start_Delay_List, A_LineNumber)
      Start_Delay_List := IniRead("Serial_Master.cfg", "Default", "Start_Time_List","ERROR")  ;if this is ERROR, usebuilt-in File_name
      if ((Start_Delay_List= "ERROR") or (Start_Delay_List= "" ))
        Start_Delay_List:=BI_Start_Delay_List                                ;last resort timer list
    }
    Start_Delay_array:=StrSplit(Start_Delay_List, "|")                    ; Split the Start_Time_List into an array
    Start_Delay:=Start_Delay_array[profile_array[8]]                     ; get profile choice Start_Delay

    ;MsgBox "Partner_app " Partner_app "`n App_exe " App_exe "`n App_profile " App_profile "`nStart_opt " Start_opt "`nFile_name " File_name "`n start timer " Start_Delay "`n line number " A_LineNumber
    ;**************************************************************************************************************************
    ;**** Application working parameters are set: Partner program named, Companion File_Name formed                  **********
    ;****  and Program Start option set.                                                                             **********
    ;**************************************************************************************************************************

    If (Partner_app ="None")  {
      Partner_ID:=0
      Partner_PID:=0
      Partner_App_Title :="None"
     return
    }
    if(Start_opt="launch") or (File_ext_array[profile_array[6]] ="none") or (File_name_array[profile_array[5]]="none") {
    Launch_Retry:
      Pgm_List := WinGetList("ahk_exe " App_exe_short,,,)  ; Check for any Ahk__exe that matches partner program. Returns an array of IDs
      ;Partner_app is not reliably found in the Partner_app Title.
      if (Pgm_List.Length =0)      {                   ; Pgm_List =0: Partner _App is not already running
        No_file_Name:=""
        Result:=Start_Application(App_exe, App_exe_short, No_file_Name,Start_Delay )
        if (Result="RUN ERROR")   {                     ; Config_Message= Failed to Launch. error message.
          ;MsgBox "launch failed  Partner_app " Partner_app  "`n  Line number " A_LineNumber
          return                                        ; failed start thread is now done & a new one may start
        }
        if (Result = "ID ERROR") {
          MsgBox("Get ID Failed " App_exe_short "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID
           . "`n Line number " A_LineNumber)
        }
        ;MsgBox(" Launch result Partner_App " App_exe_short "`n  Partner_PID " Partner_PID "`n  Partner_ID"
        ; . Partner_ID "`n  Title-" Partner_App_Title "`n  Line number " A_LineNumber)
      }
      else {                                            ; one or more instances of Partner _App is already running
        Partner_ID:="ahk_id " Pgm_List[1]               ; Grab 1st instance of Partner_app ahk_id
        Partner_PID := WinGetPID(Partner_ID)            ; Get PID number
        Partner_PID:="ahk_pid " Partner_PID
        try {
        WinWaitActive(Partner_ID, , Start_Delay)          ; Activate the old instance of Partner_app
        }
        catch {
          MsgBox(Partner_app " failed to open in " Start_Delay " Sec.`n Partner_PID-" Partner_PID
           . "`n Press OK to try again`n Line number " A_LineNumber)
          Goto Launch_Retry
        }
        else {
          ;MsgBox("Successful Re-Launch-" Partner_app "`n Partner_PID-" Partner_PID "`n Line number" A_LineNumber)
        }
      }
      Goto File_Start_Done
    }   ;End Launch process. All remaining actions are Launch + File_name
    if (!FileExist(File_Name)) and (Start_opt ="Launch, Open File") {  ;Requested file exist?
      If (InStr(File_Name, ".xl"))   {              ; Does not exist, Excel files require format info
        Try{
           FileCopy("Seed.xlsx", File_Name, 1)      ; Create a blank from seed file. (1=overwrite if it exists)
           ErrorLevel := 0
        }
        Catch as error {
          MsgBox(" FileCopy error `n what: " error.what "`nfile: " error.file
        . "`nline: " error.line "`nmessage: " error.message "`nextra: " error.extra,, 16)
        }
      }
      else
        FileAppend "", File_Name                    ; Create, then close an empty file, in this program's directory
    }   ;Requested file now exists...

    if (Start_opt ="Launch, Open File")      {
      Save_Match_Mode:=A_TitleMatchMode
      SetTitleMatchMode(2)                          ; A window's title be anywhere within the specified WinTitle to be a match.
      Name_List := WinGetList(File_Name,,,)         ; This file name in any Title?  Name_List is an array of ahk_id
      ; MatchMode(2)is used since a valid File_Name match may have one pre-pended symbol (*)
      ; Some Apps do not put file name in WinTitle (e.g. Windows Notepad 11.2302.16.0) and will not appear in Name_List.
      ; Notepad 11.2303.40.0 puts the active tab title in the WinTitle. Other tab names will not be in the list
      ; Legacy Notepad and Notepad 11.2303.40.0 will prepend the title with '*' to indicate file has changed since opened.

      SetTitleMatchMode(Save_Match_Mode)
      ;MsgBox("Name_List length " Name_List.Length  "`nName_List[]1 " 0+Name_List[1] "`nName_List[2] " 0+Name_List[2]
      ; . "`nName_List[3] " 0+Name_List[3] "`n Line number " A_LineNumber)
      Old_owner_id:=0
      if (Name_List.Length !=0) {                   ; Check if requested file is already open in any App
        Partner_ID:=0
        Loop Name_List.Length                       ; First, search list for new Partner_app.
        {
          Name_posn:=0
          Test_ID:="ahk_id " Name_List[A_Index]
          Test_PID :="ahk_pid " WinGetPID(Test_ID)
          Test_App := WinGetProcessName(Test_ID)
          Test_Title := WinGetTitle(Test_ID)        ; A valid File_Name match may have one pre-pended symbol (*)
          Test_Title := StrReplace(Test_Title, "*") ; remove it if it exists
          Name_posn:=InStr(Test_Title, File_Name)   ; Where is File_Name within the test title?
          if (Name_posn = 1) {                      ; Some App owns this file...
            ;MsgBox(File_Name " Owner found -- " Test_App "`n App Title " Test_Title "`n Line number " A_LineNumber)
            if (Test_App = App_exe_short)   {       ; Partner_App already owns this file name?
              Partner_PID:=Test_PID                 ; yes, get credentials and go
              Partner_ID:=Test_ID
              break
            }
            if (Old_owner_id =0)  {                 ; first found?
              Old_owner_id:=Test_ID                 ; First encountered owner, there could be others open.
              Old_owner_pid :=Test_PID              ; including partner app
              Old_owner:=Test_App                   ; so continue search
            }
          }
        }
        if (Partner_ID!=0)  {                       ;searched and found partner app in list
           ; MsgBox(" Re-Launch-open " Partner_app "`n Partner_ID " Partner_ID "`nPartner_PID-" Partner_PID
           ;  . "`n Line number " A_LineNumber)
            Goto File_Start_Done                      ; Activate already open Partner_app - File name combination
        }
      }
      if (Old_owner_id !=0) {                       ; list length was =0 or a name match was never found.
        SetTimer(ChangeButtonNames,10)              ; Hack to change message box button names
        ;                                           ; YES/NO/Cancel changed to Continue/Restore/Cancel
        msgResult := MsgBox(File_Name " is already open in " Old_owner ".`n Click Continue to open - " Partner_app
         . "`n            Restore to activate " Old_owner "`n            Cancel to open Program Options", "File Open Warning",19)
           ;YES/NO/Cancel Box changed to to Continue/Restore/Cancel
        if (msgResult = "NO")                       ; Restore -- Activate existing app
        {
           Partner_PID:=Old_owner_pid
           Partner_ID:=Old_owner_id
           Partner_App_Title := WinGetTitle(Partner_ID)
           WinActivate(Partner_ID)                  ; No delay parameter here, Partner_App + File_name already exist.
           ;MsgBox("restore file open " Partner_app "`n Partner_PID-" Partner_PID "`n Partner_ID-" Partner_ID
           ; . "`n Line number "A_LineNumber)
           Goto File_Start_Done                    ; Activate already open Partner_app - File name combination
        }
        if (msgResult = "Cancel")
        {
          MenuHandler()                             ; Abort this Partner App selection and go back to Program Options
          WinWaitClose("Serial Master")             ; Wait (forever) for GUI to close. Serial_Master.cfg may be new...
          return                                    ; on close, this effort is done,
        }
        if (msgResult = "YES")                      ;Continue, launch new App-Filename
        {
          Result:=Start_Application(App_exe, App_exe_short, file_Name,Start_Delay )
          if (Result="RUN ERROR")   {               ; Config_Message= Failed to Launch. error message
            return                                  ; failed start thread is now done & a new one may start
          }
          if (Result = "ID ERROR") {
            MsgBox("Get ID Failed " App_exe_short "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID "`n Line number " A_LineNumber)
          }
          ;Partner_App_Title := WinGetTitle(Partner_ID)       ; Get Application title (Not used any more?)
          ;MsgBox("Continue-Open " Partner_app "`n  Partner_ID " Partner_ID "`n  Partner_PID-" Partner_PID
          ;. "`n  Title-" Partner_App_Title "`n  Line number " A_LineNumber)
          Goto File_Start_Done
        }
      }
      else {                                        ;File is not open in any app. Good to open App with file
        Result:=Start_Application(App_exe, App_exe_short, file_Name,Start_Delay )
        if (Result="RUN ERROR")   {                         ; Config_Message= Failed to Launch. error message
          return                                            ; failed start thread is now done & a new one started
        }
        if (Result = "ID ERROR") {
          MsgBox("Get ID Failed " App_exe_short "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID,  "Master Menu Line number " A_LineNumber)
        }
        Partner_App_Title := WinGetTitle(Partner_ID)
        ;MsgBox("Clean-Open " Partner_app "`n  Partner_ID " Partner_ID "`n  Partner_PID-" Partner_PID
        ;. "`n Partner_App_Title-" Partner_App_Title "`n  Line number " A_LineNumber)
        ; Some Apps do not put file name in WinTitle (e.g. Windows Notepad 11.2302.16.0)
        ; However Partner_ID is a the definitive identifier for the window Just opened
      } ; An application should now be opened with an associated file
    }
    else    ;Wait for user to open desired Program and file. This version requires a partner program selection other then none.
    {
      MsgBox("Serial Master Waiting", "Waiting for " Partner_app " to start (and data file if needed)`n Click OK when ready to proceed.",4096 + 48)
      Open_retry:
      App_Count := WinGetCount("ahk_exe" App_exe_short)
      if (App_Count =0 ) {
        MsgBox( Partner_App " is not active, Click OK when " Partner_App " is active.", "Master Menu Error ! Line number " A_LineNumber, 4096 + 48)
        Goto Open_retry
      }
      if (App_Count >1 ) {
        MsgBox("Multiple (" App_Count ") copies of " Partner_app " are Open`n Click OK when desired Partner Program is in Focus.","Serial Master Caution !",4096 + 48 )
      }
      Partner_ID := WinGetID("ahk_exe" App_exe_short)
      Partner_ID:="ahk_id " 0+Partner_ID
      Test_PID := WinGetPID(Partner_ID)
      Partner_PID:="ahk_pid " Test_PID
      ;MsgBox("Wait-User " Partner_app "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID "`n Line number " A_LineNumber)
    }

  File_Start_Done:
    Partner_App_Title := WinGetTitle(Partner_ID)       ; Get Application title (Not used any more?)
    ;MsgBox ("Partner_App-" Partner_App "`nPartner_App_Title-" Partner_App_Title "`nPartner_PID-" Partner_PID "`nPartner_ID-" Partner_ID "`n Line number " A_LineNumber)
    WinActivate(Partner_ID)
    Send("^{END}")                                        ;set cursor to end of file
    return
  }

  Start_Application(App_exe, App_exe_short, File_Name,Start_Delay_x)   {   ; Starts an application program (with File_Name) and verifies application is up and running

    global  Partner_ID:=0
    global  Partner_PID:=0
    global  Partner_App
    global  Config_Message
    global  Config_Font
    Exit_code:=0
    try {
      Run(App_exe " " File_Name, , "", &Partner_PID)    ;Open Application with File_Name
    }
    catch as error {                      ;Program launch failed...
      ;MsgBox(" run command error `n what: " error.what "`nfile: " error.file
      ;  . "`nline: " error.line "`nmessage: " error.message "`nextra: " error.extra,, 16)
      Config_Message:= Partner_App " Failed to Launch. " App_exe_short " `n Error message: " error.extra
       . " Choose another Progran and click OK, or`n Click CANCEL to Default to None, or EXIT"
      Config_Font:="s10 cred"
      Partner_App := "None"                             ; This will keep Rx Loop from showing error
      MenuHandler()                                     ; Partner program is now 'None' let user figure out what to do.
      WinWaitClose("Serial Master")                     ; Wait (forever) for GUI to close. Serial_Master.cfg may be new...
      Exit_code:= "RUN ERROR"                           ; MenuHandler() Cancel should return here with no changes
      ;MsgBox " Return from MenuHandler() Close. Partner_App " Partner_App "`nPartner_ID " Partner_ID "`n Line number " A_LineNumber
      return   (Exit_code)                              ;this should end the failed run attempt thread
    }
    else {        ; No Run Error
      Partner_PID:="ahk_pid " Partner_PID
      Partner_ID := WinWaitActive(Partner_PID, , Start_Delay_x) ;returns Partner_ID if app actually becoms Active, 0 if timeout.
      if ( Partner_ID=0)     {                          ; PID=0 is timeout and app has not started
        if (Start_Delay_x <= 3) {                       ; If an application starts out with a splash or any transient display window
          Added_delay:=3-Start_Delay_x                  ; that lasts for more then Start_Delay_x AND has a PID or ID that differs from
                                                        ; the working window, this process may capture the wrong PID and ID.
        }
        else {
          Added_delay:=.5                                 ; Wait at leasst 3 sec for any transient/ Splash to clear
        }
        Partner_ID := WinWaitActive(Partner_PID, , Added_delay) ;returns Partner_ID if app actually becoms Active, 0 if timeout.
        if ( Partner_ID=0)     {                        ; Windows 11 Note pad 11.2302.16.0 does not start with a 'Working" PID
          ; MsgBox("WinWaitActive Failed " App_exe_short "`n Partner_PID-" Partner_PID "`n Line number " A_LineNumber)
          try {
            Partner_ID := WinGetID("ahk_exe " App_exe_short) ; This should be the one just opened. Notepad can have more then 1 ID here.
          }
          catch { ; A TargetError is thrown if the window could not be found
            Exit_code:= "ID ERROR"
            MsgBox(" ID recovery Failed " App_exe_short "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID "`n Line number " A_LineNumber)
          }
          else {          ; ID Recovery successful
            ;Partner_ID:=0+Partner_ID                              ;hex to decimal conversion probably not necessary in v2
            Partner_ID:="ahk_id " Partner_ID
            Partner_PID := WinGetPID(Partner_ID)                 ; Get the correct PID
            Partner_PID:="ahk_pid " Partner_PID                  ;Credentials should be good to go
            ; MsgBox(" Start_Application " App_exe " ID Recovery successful`n Partner_PID " Partner_PID "`n Partner_ID " Partner_ID "`n Line number " A_LineNumber)
          }
        }
        else {  ;clean WinWaitActive 2nd try
          Partner_ID:="ahk_id " Partner_ID
          ; MsgBox(" Clean Application Start (2nd try) " App_exe "`n Partner_PID-" Partner_PID "`n Partner_ID " Partner_ID "`n Line number " A_LineNumber)
        }
      }
      else {  ;clean WinWaitActive 1st try
        Partner_ID:="ahk_id " Partner_ID
        ; MsgBox(" Clean Application Start (1st try) " App_exe "`n Partner_PID-" Partner_PID "`n Partner_ID " Partner_ID "`n Line number " A_LineNumber)
      }
    }
    if (Exit_code =0) {
      Partner_App_Title := WinGetTitle(Partner_ID)      ; Get title from Application.
      ;MsgBox(" Retrieved Title-" Partner_App_Title "`nLine number " A_LineNumber)
      if (Partner_App_Title !="") and (File_Name !="") and (InStr(Partner_App_Title, File_Name)!=1) {
        Exit_code :="Title Error"       ; This window's title must start with the specified File_Name to be a match.
        ToolTip("Close Pop-up window to continue", 100, 50) ; No match, This ID does not have a 'clean' Title
        WinWaitClose(Partner_ID)       ; Application needs user input to finish open. e.g. Winword file conversion
        ToolTip()        ;next recover the ID
        try {
        Partner_ID := WinGetID("ahk_exe " App_exe_short)  ; This should be the one just opened. Notepad can have more then 1 ID here.
        }
        catch { ; A TargetError is thrown if the window could not be found
          Exit_code:= "ID ERROR"
          ;MsgBox(" Get ID Failed " App_exe_short "`n Partner_ID " Partner_ID "`n Partner_PID-" Partner_PID "`n Line number " A_LineNumber)
        }
        else {                                            ; ID Recovery successful, get PID and clear error code
          Partner_ID:="ahk_id " Partner_ID
          Partner_PID := WinGetPID(Partner_ID)
          Partner_PID:="ahk_pid " Partner_PID                    ;Credentials should be good to go
          Exit_code :=0
          ;MsgBox (" Start_Application " App_exe " Title Error Recovery-`n Partner_PID-" Partner_PID "`n Partner_ID " Partner_ID "`n Line number " A_LineNumber)
        }
      }
    }
    Return (Exit_code)
  }

ChangeButtonNames() {
  if !WinExist("File Open Warning") {
     return  ; Keep waiting for MsgBox named "File Open Warning".
  }
  SetTimer(,0) ;kill the timer
  WinActivate()
  ControlSetText("&Continue", "Button1")
  ControlSetText("&Restore", "Button2")
return
}
