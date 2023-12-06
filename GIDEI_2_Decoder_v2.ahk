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
#Requires AutoHotkey v2+

;########################################################################################################
;################      GIDEI Escape sequence Decoder  Components     ####################################
;#################       Called by the   RS232 COM port receive process         #########################
;########################################################################################################
; GIDEI_2_Decoder_V2.ahk  ... fixed conversion issues and all commands retested
; Edited 07/17/2023 Written and edited by HV Paul
;          Com_Config_GUI(Com_inifile) was  Com_Config_GUI()
;          Added Decoder_init()  that declares Key_List Global
; This module has two main parts:
; 1)  The Decode_serial(Byte) function captures escape sequences, completes "On Hold" escape sequences and
; calls the main decoder process.The Decode_serial(Byte) function must be called by the serial data input
; process with a single character code in hexidecimal format (0xNN) .   Decode_serial(Byte) returns serial
; input data that is not part of an escape sequence or returns a empty string if input is used by the
; process. Decode_serial(Byte) may #include Test_kit.ahk which is used for insertion or
; detection of code test sequences.
; 2) Send_esc(Esc_Sequence) is the GIDEI (2.1) Standatd Escape Sequence Decoder.
; ******** The decoder expects the following resources to be available *********************************
; **** RS232_Write(RS232_FileHandle,0x11) function is called to send Xon in response to NULL.
; **** Subroutine Com_Config_GUI(Com_inifile)  Called to request Baudrate change using global Com_Config_GUI

;Global Com_inifile                 ; Com_inifile is global Read Only

Decoder_init() {                    ; Call this in autoexec section to prepare for operation
CoordMode "Mouse", "Screen"         ; use full screen coordinates for mouse
Global Key_List:=("alt.appskey.backspace.BS.capslock.ctrl.del.delete.down.end.enter.esc.escape.f1.f2.f3.f4.f5.f6.f7.f8.f9.f10."
. "f11.f12.f13.f14.f15.f16.f17.f18.f19.f20.f21.f22.f23.f24.home.ins.Insert.LAlt.LButton.Lcontrol.LCtrl.left.LShift.LWin.MButton."
. "numlock.pause.pgdn.pgup.printscreen.ralt.RButton.rcontrol.rctrl.right.rshift.RWin.scrolllock.shift.Sleep.Space.tab.up.WheelDown."
. "WheelLeft.WheelRight.WheelUp.NumLock.Numpad0.Numpad1.Numpad2.Numpad3.Numpad4.Numpad5.Numpad6.Numpad7.Numpad8.Numpad9.NumpadAdd."
. "NumpadClear.NumpadDel.NumpadDiv.NumpadDot.NumpadDown.NumpadEnd.NumpadEnter.NumpadHome.NumpadIns.NumpadLeft.NumpadMult."
. "NumpadPgDn.NumpadPgUp.NumpadRight.NumpadSub.NumpadUp.")
;msgbox ("This is GIDEI_2_Decoder_V2.ahk: AutoHotKey version" A_AhkVersion "`n Line number " A_LineNumber)
return
}

Decode_serial(Byte) {
  ; Byte is the hex value of input character code and will be consumed if it is part of the escape
  ; sequence assembly process.  If byte is not part of an escape sequence, the input character code is returned.
  ;
static Esc_Sequence :=""            ; these are initialized only once...
static Hold_modifier :=""
 ;#Include Test_Kit_V2_click_dblclick_test.ahk  ;only for code test
  if (byte= 0x00) {                 ; received NULL
    RS232_Write(RS232_FileHandle,0x11) ;Send Xon out the RS232 COM port
  }
  ; Note that for each byte processed, only one of the 4 if/else if/else conditions runs
  if (byte= 0x1b) {                 ; escape char ALWAYS starts a new sequence
    Esc_Sequence :="<esc>"          ;and prior escape nor modifiers are retained
    Hold_modifier:=""               ; New modifiers(s) may be returned by an escape sequence
    byte:=""                        ; Byte is 'consumed'
  }
  else if (Esc_Sequence != "") {    ;continue to build on the sequence
    Esc_Sequence .=Chr(Byte)
    if (Byte = 0x2e) {              ;decimal point always ends an escape sequence
      Hold_modifier :=Send_esc(Esc_Sequence)    ;process escape command, catch any returned hold modifiers
      Esc_Sequence :=""
    }
    byte:=""                        ; Byte is 'consumed'
  }
  else if Hold_modifier        {    ; not executed if either of the above are...
    Hold_modifier .=Chr(Byte)       ; append input to any previously aquired 'Hold" result
    if (SubStr(Hold_modifier, 1, 5) != "<esc>")  {
      Send(Hold_modifier)         ; send modifier and the single key we were waiting for.
      Hold_modifier :=Send_esc("<esc>,rel,No_key.") ;dummy release command to clear modifiers in escape decoder
    }
    else {                         ; complete held escape sequence and process it
      Hold_modifier .="."
      Hold_modifier :=Send_esc(Hold_modifier)    ;process escape command, catch any returned hold modifiers
    }
    byte:=""                        ; Byte is 'consumed'
  }
return byte
}

;*********** End Decode_serial(Byte)******************************************************************************************
Send_esc(esc_seq) {
  ;Submitted escape sequence can be a command, command modifiers, key name(s), Character_Name, short-cut and may contain
  ;alias(es) Esc_modifier may be returned as any residual 'hold' data
  ;
  ; Character_Names variable holds key names that have no symbol within ASCII character code points 32-127 and are rendered
  ; as Unicode characters. Characters that have an upper case symbol are prefixed with a '+'.  The string also includes
  ; the special cases of comma, plus, period and pound since they have special meaning in escape sequences and can't be treated
  ; as aliases. Shift,comma ('<') and shift,period ('>') are also included.
Global Config_Message
Global Config_Font

    static Character_Names := "`n+aacute.00C1`naacute.00E1`n+acircumflex.00C2`nacircumflex.00E2`nacute.00B4`n+adieresis.00C4`n	adieresis.00E4`n+ae.00C6`nae.00E6`n+agrave.00C0`nagrave.00E0`n+aogonek.0104`naogonek.0105`n+aring.00C5`naring.00E5`nbbar.00A6`n+cacute.0106`ncacute.0107`n+ccaron.010C`nccaron.010D`n+ccedilla.00C7`nccedilla.00E7`ncedilla.00B8`ncircumflex.02C6`ncomma.002C`n+comma.0036C`ndieresis.00A8`ndivide.00F7`n+eacute.00C9`neacute.00E9`n+ecaron.011A`necaron.011B`n+ecircumflex.00CA`necircumflex.00EA`n+edieresis.00CB`nedieresis.00EB`n+egrave.00C8`negrave.00E8`n+eth.00D0`neth.00F0`nexclaimdown.00A1`ngrave.0060`n+iacute.00CD`niacute.00ED`n+icaron.01CF`nicaron.01D0`n+icircumflex.00CE`nicircumflex.00EE`n+idieresis.00CF`nidieresis.00EF`n+igrave.00CC`nigrave.00EC`nmicro.00B5`nmordinal.00BA`nmultiply.00D7`n+ncaron.0147`nncaron.0148`n+ntilde.00D1`nntilde.00F1`n+oacute.00D3`noacute.00F3`n+ocircumflex.00D4`nocircumflex.00F4`n+odieresis.00D6`nodieresis.00F6`n+oe.0152`noe.0153`n+ogonek.01EA`nogonek.01EB`n+ograve.00D2`nograve.00F2`n+ohungarumlaut.0150`nohungarumlaut.0151`nonehalf.00BD`nonequarter.00BC`n+oblique.00D8`nooblique.00F8`n+otilde.00D5`notilde.00F5`nperiod.002E`n+period.003E`nplus.002B`npound.0023`nrcaron.0159`n+rcaron.0158`nring.02DA`n+sacute.015A`nsacute.015B`n+scaron.0160`nscaron.0161`nsection.00A7`nsharps.00DF`nsuperone.00B9`nsuperthree.00B3`nsupertwo.00B2`n+tcaron.0164`ntcaron.0165`ntilde.007E`n+uacute.00DA`nuacute.00FA`n+ucircumflex.00DB`nucircumflex.00FB`n+udieresis.00DC`nudieresis.00FC`n+ugrave.00D9`nugrave.00F9`n+uhungarumlaut.0170`nuhungarumlaut.0171`n+uring.016E`nuring.016F`n+yacute.00DD`nyacute.00FD`n+ydieresis.0178`nydieresis.00FF`nyen.00A5`nZcaron.017D`nzcaron.017E`n+zdotaccent.017B`nzdotaccent.017C"

static alias_list:="`naltgr=ctrl,alt`namp=&`nampersand=&`napostrophe='`nast=*`nasterisk=*`nat=@`nbackslash=\`nbreak=pause`nbslash=\`nbspace=bs`nbut1=Lbutton`nbut2=Mbutton`nbut3=Rbutton`nbut4=XButton1`nbut5=XButton2`nbutdefault=Lbutton`ncancel=esc`ncapslk=CapsLock`nclear=ctrl,x`ncolon=:`ncopy=ctrl,c`ncontrol=ctrl`ncut=ctrl,x`ndblquote=u+0022`ndn=down`ndollar=$`neight=8`nequal=U+003D`nexclaim=u+0021`nfive=5`nfour=4`nfullsize=win,up`nhelp=f1`nhyphen=-`nkp-=NumpadSub`nkp*=NumpadMult`nkp/=NumpadDiv`nkp+=NumpadAdd`nkp0=Numpad0`nkp1=Numpad1`nkp2=Numpad2`nkp3=Numpad3`nkp4=Numpad4`nkp5=Numpad5`nkp6=Numpad6`nkp7=Numpad7`nkp8=Numpad8`nkp9=Numpad9`nkpdel=NumpadDel`nkpdivide=NumpadDiv`nkpdn=NumpadDown`nkpdown=NumpadDown`nkpdp=NumpadDot`nkpend=NumpadEnd`nkpenter=NumpadEnter`nkphome=NumpadHome`nphyphen=NumpadSub`nkpins=NumpadIns`nkpinsert=NumpadIns`nkpleft=NumpadLeft`npminus=NumpadSub`nkpperiod=NumpadDot`nkppgdn=NumpadPgDn`nkppgup=NumpadPgUp`nkpplus=NumpadAdd`nkpright=NumpadRight`nkpslash=NumpadDiv`nkpstar=NumpadMult`nkptimes=NumpadMult`nkpup=NumpadUp`nmenu=appskey`nminus=-`nnext=pgup`nnine=9`nnumber=NumLock`nnumlk=NumLock`none=1`npagedown=pgdn`npageup=pgup`npaste=ctrl,v`nplus=u+002B`npound=u+0023`nprev=pgup`nprint=^p`nprtscr=PrintScreen`nrbrace=}`nrbracket=]`nrcompose=compose`nremove=del`nret=enter`nreturn=enter`nrightwinkey=RWin`nrolldown=WheelDown`nrollup=Wheelup`nrparen=)`nsemicolon=;`nseven=7`nshiftleft=Shift`nshiftright=Shift`nsix=6`nslash=/`nthree=3`ntwo=2`nunderscore=_`nundo=ctrl,z`nwin=Lwin`nzero=0`n"



    static Mouse_keys:="LButton.RButton.MButton.XButton1.XButton2."    ;any mouse buttom must use these names or an alias that equates to one
    static M_dir:=".up=+0.-1.down=+0.+1.left=-1.+0.right=+1.+0.upleft=-1.-1.upright=+1.-1.downleft=-1.+1.downright=+1.+1."
    static Modifiers:=""
    static Lock_list:=""
    static M_Lock_list:="."
    static anchor_list:=""
    static Mouse_Stop:=1                            ; stop=1, mouse movement has stopped
    static Pix2SysX:=65535 / A_ScreenWidth          ; constants to convert 1 pixel to system desktop units
    static Pix2SysY := 65535 / A_ScreenHeight	    ; used in mouse_move_Loop:
    static kx
    static ky
    static loop_time
    static StartTime                                ; time mouse move started (diagnostic)
    static x1
    static y1
    static s_field
    esc_seq :=StrReplace(esc_seq, A_Space)          ; squeeze out any spaces
    esc_seq := StrLower(esc_seq)                    ; All lower case here
    esc_sav:=esc_seq                                ; save for failed process
    fail_hint:=""                                   ; clear the failure hint
    esc_seq := SubStr(esc_seq, 6)                   ; remove '<esc>'
    esc_seq :=StrReplace(esc_seq, ",", ".")         ; all fields now end in '.'
    ; Do alias replace HERE on all fields
    Loop Parse esc_seq, "."                         ;extract escape fields and check for an alias
    {
        if (alias_start:=InStr(alias_list, "`n" . A_LoopField . "="))  { ;find the start of an exact alias match

          alias_start:=InStr(alias_list, "=", , alias_start)+1 ;move up passed the'='
          alias_end:=InStr(alias_list, "`n", , alias_start) ;find the end of match
          esc_seq:= StrReplace(esc_seq, A_LoopField, SubStr(alias_list, alias_start, alias_end-alias_start))
          esc_seq :=StrReplace(esc_seq, ",", ".")         ; replace any "," with "."
        }  ;no match, don't touch the field
    }
    cmd_str:=""                                     ; clear old command, if any
    if (SubStr(esc_seq, 1, 1) =".")  {              ; separate new command, if any
      cmd_str:=SubStr(esc_seq, 1, InStr(esc_seq, ".", , 2)) ;cmd_str is from 1st chr up to and includes the 2nd '.'
      esc_seq :=StrReplace(esc_seq, cmd_str)        ;remove command from esc_seq
    }
    if (InStr(esc_seq, "."))     {                  ;esc_seq can be empty or have at least 1 remaining '.'
      esc_seq :=SubStr(esc_seq, 1, -1)                ;remove last '.'
    }
   ;MsgBox ("escape command=" cmd_str "`nescape fields " esc_seq "   `n Line number " A_LineNumber)
    ;Complex GIDEI commands 'lock'and''rel' treat modifiers as key names. Note that any field contents,
    ;  valid or not, can be locked or released.
    Switch cmd_str                                 ; First process commands that use only key names
    {
      Case ".rel." :
      {
        if !(StrLen(esc_seq))   {                   ;Release All if no key fields
         SetCapsLockState("Off")                     ; CapsLock does not toggle with key press...
         Loop 255
          {
           if GetKeyState(Key:=Format("VK{:X}",A_Index))  {
             SendInput("{" Key " up}")                  ;Release if down and valid
           }
          }
          Lock_list:=""                             ; clear Lock_list
        }  ;end release all
        else {                                   ; release command with key list
         Loop Parse esc_seq, "."                  ;extract and release Key_names
         {                                          ; Only case insensitive but otherwise exact match keynames are released.
           if (InStr(Lock_list, A_LoopField))  {
             Lock_list:=StrReplace(Lock_list,  A_LoopField . ".") ; remove 'A_LoopField.' from list
             SendInput("{" A_LoopField " up}")          ;Release it.
            }
          }
        }  ;end release list
        ;MsgBox (cmd_str "command " esc_seq "`n locked list=" Lock_list  "`n Line number " A_LineNumber)
        Modifiers:=""                               ;clear modifiers even if no key is released
        return Modifiers
      }    ; end .rel. case
      Case ".lock.":
      {
        Loop Parse esc_seq, "."                     ;extract Key_names
        {                                           ; locked keys ar stored as '.key.' to prevent aliasing in search
          SendInput("{" A_LoopField " down}")       ;lock it down.  MAY LOOK LIKE A KEY PRESS!
          Lock_list:=StrReplace(Lock_list, A_LoopField  ".") ; If exists, remove A_LoopField from anywhere in list
          Lock_list.=A_LoopField . "."              ; put new lock key at end of the list
        }
        ;MsgBox (cmd_str "command " esc_seq "`n locked list=" Lock_list  "`n Line number " A_LineNumber)
        Modifiers:=""
        return Modifiers
      }  ;end .lock. command  case

      Default:                                     ;The remaining commansds allow modifiers... They are collected here
      {
        esc_seq.="."                                ;restore the last '.' for modifiers processing
        if (InStr(esc_seq, "shift.")) {              ; (Any) shift. ?
          esc_seq:=StrReplace(esc_seq, "shift.")     ;remove any shift from esc_seq
          Modifiers:=StrReplace(Modifiers, "+")      ;remove any shift from modifiers
          Modifiers:="+" Modifiers                ;prepend a shift modifier
        }
        if (InStr(esc_seq, "ctrl.")) {               ; (Any) ctrl. ?
          esc_seq:=StrReplace(esc_seq, "ctrl.")
          Modifiers:=StrReplace(Modifiers, "^")
          Modifiers:="^" Modifiers                ;prepend a conrtol modifier
        }
        if (InStr(esc_seq, "alt.")) {                ; (Any) alt. ?
          esc_seq:=StrReplace(esc_seq, "alt.")
          Modifiers:=StrReplace(Modifiers, "!")
          Modifiers:="!" Modifiers                ;prepend a alt modifier
        }
        if (InStr(esc_seq, "Lwin.")) {               ; (Any) "win.. ?
          esc_seq:=StrReplace(esc_seq, "Lwin.")
          Modifiers:=StrReplace(Modifiers, "#")
          Modifiers:="#" Modifiers                ;prepend a win modifier
        }
        if ((esc_seq ="") and ( cmd_str ="")) {
         ;MsgBox ("implied  modifiers ="Modifiers )
          return Modifiers                         ;treat as implied 'modifier hold' and return
        }
        esc_seq :=SubStr(esc_seq, 1, -1)              ;remove last '.'
      }
    }       ;End name list Switch

    ;MsgBox ("escape sequence=" cmd_str esc_seq "   modifiers =" Modifiers "  `n Line number " A_LineNumber )
    ;At this point, <esc> has been removed, Key modifiers have been moved to Modifiers, any command is moved to cmd_str and
    ; esc_seq can be empty
    Switch cmd_str  {                               ; Process commands that use key modifiers
      Case ".hold." :
      {                            ; esc_seq should not be empty, but it is not an error
        Modifiers:= Modifiers esc_seq               ;save hold data AFTER any modifiers
        ;MsgBox (".hold. modifiers =" Modifiers )
         return Modifiers
      }
      Case ".combine." :
      {                         ; should have only one field remaining....
        esc_seq:=StrReplace(esc_seq, ".") ; Concatinate any remaining fields and process as implied press
      }
      Case ".baudrate." :                          ; typically will have one field remaining....
      {
        Config_Message:=" GIDEI Decoder requests Bit Rate " esc_seq " baud"
        Config_Font:= "cblue"                        ; Config_Font is Blue
        Com_Config_GUI(Com_inifile)
        Modifiers:=""
        return Modifiers
      }
      Case ".anchor." :
      {
          if (esc_seq ="")  {                       ; Anchor with no assigned letter is a "hold <esc> for assignment"
            Modifiers:="<esc>,anchor,"              ; return this command as an escape sequence and next a-z input will
            return Modifiers                        ; be appended and re-processed to establish an anchor point
          }                                         ; next character typed returns with "<esc>,anchor,X.
          else if (strlen(esc_seq) =1)  {           ; looks like command is complete either by 'hold' or
            if isAlpha(esc_seq)                     ; letter supplied with anchor(not prohibited by standard)
            {
              if (anchor_x:=InStr(anchor_list, esc_seq)) { ;check if esc_seq is in anchor list
                old_anchor:=SubStr(anchor_list, anchor_x, InStr(anchor_list, ".", , anchor_x, 2))
                anchor_list:=StrReplace(anchor_list, old_anchor)  ; old anchor is removed from list
              }  ;else/then add to list
              MouseGetPos(&X_coord, &y_coord)         ; get cursor X,Y position
              anchor_list.=esc_seq "=" X_coord "."  Y_coord "." ;append esc_seq and X,Y to anchor_list
              ;MsgBox ("valid anchor command -" cmd_str esc_seq "`n anchor list=" anchor_list )
              Modifiers:=""
              return Modifiers
            }
          }
          fail_hint:="`n Invalid Anchor Lablel"
          Goto Bad_code
      } ;end .anchor. command

      Case ".goto." :                                   ; This command moves the mouse cursor to Coordinates relative to the
      {                                                 ; full screen upper left corner.
                                                        ; .goto. command has been removed from esc_seq
        Mouse_Stop:=1                                   ; Stop any movement due to Mougo   goto with no coordinates
          if (esc_seq ="")  {                           ; nor a-z lable is a "hold <esc> for a-z assignment"
            Modifiers:="<esc>" cmd_str                ; return this command as an escape sequence and next a-z input will
            return Modifiers                            ; be appended and re-processed to determine the a-x goto point
          }
        X_Y_str:=esc_seq                                ; save as possible numeric X,Y
        if (strlen(esc_seq) =1)  {
          if isAlpha(esc_seq)                           ;check goto with single a-z lable
          {
            if (list_posn:=InStr(anchor_list, esc_seq)) { ;find goto lable in anchor list then clip up to and including lable's =
              X_Y_str:=SubStr(anchor_list, list_posn+2)  ;next two values in anchor_list are this lable's X,Y
            } ;X_Y_str is now either the .goto. field(s) X & Y string(s)  or X & Y coordinates from anchor_list.
          }
        }
        ;Using 'Screen' relative coordinates, both X & Y must exist as positive numbers
        X_coord:=""
        Y_coord:=""
        Loop Parse X_Y_str, "."                      ; verify that X_Y_str has X,Y coordinates
        {
          if (A_Index =1)
            X_coord:=A_LoopField
          else if (A_Index =2)
            Y_coord:=A_LoopField
          else break
        }
        ;MsgBox ("X_coord is " X_coord " Y_coord is " Y_coord "`n Line number " A_LineNumber)
        if( IsInteger(X_coord)  and X_coord >=0 and IsInteger(Y_coord) and Y_coord >=0) {
          MouseMove(X_coord, Y_coord)                   ;only real positive numbers allowed
          Modifiers:=""
          return Modifiers
        }                                               ;otherwise bad .goto. command
        fail_hint:="`n Invalid or Unknown Goto co-ordinates"
        Goto Bad_code
      } ;end .goto. case

      Case ".move." :                                  ;  .move. command has been removed from esc_seq
      { ;NOTE changing the speed parameter has no effect here
        Mouse_Stop:=1                                   ; Stop any movement due to Mougo
        X_Y_str:=esc_seq                                ; save X,Y, S
        X_coord:=0
        Y_coord:=0
        Loop Parse X_Y_str, "."                      ; verify that esc_seq has X,Y coordinates
        {
          if (A_Index =1)
            X_coord:=A_LoopField
          else if (A_Index =2)
            Y_coord:=A_LoopField
          else break
        }
        if isInteger(X_coord) and IsInteger(Y_coord)
        {
            MouseMove(X_coord, Y_coord, 0, "R")    ;mousemove relative
            ;MsgBox (cmd_str " X=" X_coord " y=" Y_coord " successful")
            Modifiers:=""
            return Modifiers
        }   ;move fails if coordinates are not integers
        fail_hint:="`n Invalid Move parameter"
        Goto Bad_code
        ;MsgBox(" invalid .move. command " esc_seq )
      } ;END .move. case

      Case ".moureset." :    ;NOTE Case ".moureset.": and  Case ".mourel." :  are not "Blocked" by{} to allow the goto
            Mouse_Stop:=1                               ; Stop any movement due to Mougo
            MouseMove(0, 0)                             ; mousemove to 'home'
            anchor_list:=""                             ; clear the anchor list
            Goto mouserel                               ;continue at mouse release with esc_seq =""
      Case ".mourel." :
      mouserel:
         if !(StrLen(esc_seq))   {                      ;Release All if no key fields
           Loop Parse Mouse_keys, "."
            {
             if (A_Loopfield !="" and GetKeyState( A_Loopfield))  {
               SendInput("{" A_Loopfield " up}")        ;Release if down and valid
             }
            }
            M_Lock_list:=""                             ; clear Mouse Lock_list
          }  ;end release all mouse buttons
         else {                                         ; release command with key list
            Loop Parse esc_seq, "."                     ;extract and release Key_names
            {                                           ; Only case insensitive but otherwise exact match keynames are released.
              Mou_But:= A_LoopField                     ;can't re-assign A_LoopField
              if (A_LoopField= "right")
                Mou_But:= "Rbutton"                     ; but must recognize right ^ left as special case...
              if (A_LoopField= "left")
                Mou_But:= "Lbutton"
              if (InStr(Mouse_keys, Mou_But "."))  {    ; make certain it is a mouse key
               M_Lock_list:=StrReplace(M_Lock_list, Mou_But ".") ; remove '.A_LoopField.' from list
               SendInput("{" Mou_But " up}")                ;Release it.
              }
              else {
               fail_hint:="`n Invalid mouse button name"
                Goto Bad_code
              }
            }
          }  ;end mourel /w list
          ;MsgBox( cmd_str " command sequence " esc_seq " `n locked mouse key list=" M_Lock_list "`n Line number " A_LineNumber)
          Modifiers:=""
          return Modifiers       ;end case
          ; no end brace this case...
      Case ".moulock." :
      {

        Loop Parse esc_seq, "."                        ;extract button names
        {        ; locked buttons ar stored as '.button.' to prevent aliasing in search
          Mou_But:= A_LoopField                         ;can't re-assign A_LoopField
          if (A_LoopField= "right")
            Mou_But:= "Rbutton"                         ; but must recognize right or left as special case...
          if (A_LoopField= "left")
            Mou_But:= "Lbutton"
          if InStr(Mouse_keys,  Mou_But ".") {              ; Check for valid button name
            SendInput("{" Mou_But " down}")                 ;lock it down
            M_Lock_list:=StrReplace(M_Lock_list, Mou_But ".") ; If exists, remove Mou_But from anywhere in list
            M_Lock_list.= Mou_But "."                 ; put new lock key at end of the list
          }
          else {
            fail_hint:="`n Invalid mouse button name"
            Goto Bad_code
          }
        }
        ;MsgBox(  cmd_str " command  " esc_seq "`n locked mouse key list=" M_Lock_list "`n Line number " A_LineNumber )
        Modifiers:=""
        return Modifiers
      }
      Case ".mougo." :           ; Slow Mouse Move command
      {
        SetTimer mouse_move_Loop, 0  		            ;stop any active timer driven move loop
        dir_field:=""
        s_field:=""
        Loop Parse esc_seq, "."                         ;extract direction and speed  from esc_seq
        {
          if (A_Index =1)                               ; first param is direction 'up', 'upright','right', 'downright' etc.
            dir_field:="." A_LoopField "="          ;bracket with '.' and '=' for exact match
          else if (A_Index =2)                          ; Mouse cursor movement is fairly slow here
            s_field:=A_LoopField                        ; 1 is slow(28sec top to bottom of screen,
          else break                                    ; 10 is faster 10 Sec. top to bottom of screen
        }
        if (dir_posn:=InStr(M_dir, dir_field))  {       ;get position in M_dir list and make certain it is valid
          dir_posn:=InStr(M_dir, "=", , dir_posn, 1)    ;mark next '='
          X_Y_str:=SubStr(M_dir, dir_posn+1, 5)         ;extract X,Y vectors after dirfield=
          if (s_field >= 1 && s_field <= 10)            ; Speed must be 1-10
          {
            MouseGetPos(&x1, &y1)						; mouse starting position
            StartTime := A_TickCount                    ; used for speed testing
            Mouse_Stop:=0                               ; stop=1, movement has stopped
            tot_time_mS:=1000 *(30 - (2*s_field))       ; convert speed to time to traverse full vertical screen
            loop_time:=30							    ; fixed time of mouse cursor step (mS)
            k:=(A_ScreenHeight*(loop_time+1.5))/tot_time_mS	; number of pixels mouse must move to meet speed
            Loop Parse X_Y_str, "."                  ; extract X, y vectors from X_Y_str
            {
              if (A_Index =1)
                kx:=A_LoopField*k                       ;kx is the direction and number of pixels/step to move
              else if (A_Index =2)
                ky:=A_LoopField*k
              else break
            }
            ;MsgBox ("mouse go direction" dir_field "  Speed" s_field "`n relative movement X=" kx " Y=" ky "`n Line number " A_LineNumber)
            SetTimer mouse_move_Loop, loop_time         ;start timer driven move loop
            Modifiers:=""
            return Modifiers       ;end done, mouse moves as a background task
          } ; else speed field failed
            fail_hint:="`n Invalid Mougo speed"
            Goto Bad_code
        } ;else direction field failed
        fail_hint:="`n Invalid Mougo direction"
        Goto Bad_code
      mouse_move_Loop() {
          x1:= x1+ kx, y1:= y1+ ky					;calculate next step then check if in bounds
          if ((x1 <= A_Screenwidth) and  (x1>=0) and (y1 <= A_ScreenHeight) and  (y1>=0) and  (Mouse_Stop=0)) {
            DllCall("mouse_event", "UInt", ABSMOVE := 32769, "UInt", x1 *Pix2SysX, "UInt", y1 * Pix2SysY)
            return
          }
          SetTimer(,0)
          Mouse_Stop:=1
          ElapsedTime :=( A_TickCount - StartTime)/1000
          MouseGetPos(&x3, &y3)
          ;MsgBox( "Result"  "speed " s_field "    time " ElapsedTime "`nloop time " loop_time "     kx=" kx "   ky=" ky "`n end X" x1 "     end Y " y1 ,,64	)
          return
        }

      }
      Case ".moustop." :
      {
        Mouse_Stop:=1
        ;MsgBox (" Mouse Stop!")
        Modifiers:=""
        return Modifiers       ;and done
      }
      Case ".click." :
      {
        N_Clicks:=1
        goto X_Click
      }
      Case ".dblclick." :
      {
        N_Clicks:=2
        goto X_Click
      }
      Default:                   ;unknown command  ?
        {
          if (cmd_str)  {                               ; Unknown command....
            fail_hint:="`n Unknown Command field"
            Goto Bad_code
          }
        }
    }       ;END switch
    ; anything else evaluate as implied press and bad values may be impossible to detect

   ;MsgBox(" implied press=" Modifiers esc_seq "`n line number " A_LineNumber)
    if (strlen(esc_seq))                {               ; No command, but something remains
      if (InStr(lock_list, "shift")) {                   ; (Any)shift lock?
        esc_seq:=StrReplace(esc_seq, "+")                ;remove any shift from esc_seq
        Modifiers:=StrReplace(Modifiers, "+")            ;remove any shift from modifiers
        Modifiers.="+"                                  ;keep only one shift as modifier
      }
                    ;NOTE Many GIDEI character codes names are not known to the system
                    ; they ALL must be looked up, translated to a Unicode key  code and sent
                    ; GIDEI specification does not require implied commands with only one field
                    ; Shift is allowed here to prefix + to a lower case character code name
                    ; prior to look-up. Any other modifier will cause the search to fail.
                    ;Character names must always be the last field of an escape sequence.
      ; First check for U_char condition. Each entry starts with "`n"
      if (Char_posn:=InStr(Character_Names, "`n" Modifiers esc_seq "."))  { ;character name search (exact match with any modifiers)
        Char_posn:=InStr(Character_Names, ".", false, Char_posn)  ;find the  associated '.'
        esc_seq  :=  SubStr(Character_Names, Char_posn+1, 4) ;get the 4 Hex character unicode number
        ;MsgBox(" U_char name found`n U+" esc_seq)
        Send("{U+" esc_seq "}")                             ;send unicode character,done with this field
        Modifiers:= ""                                    ;clear after successful character replacement
        return Modifiers                               ;clear after control sent
      }
      else if (InStr(Key_List, ("." esc_seq ".")))  { ;Search for System key name,exact match. (w/o Modifiers).
        ;MsgBox( esc_seq " is a named key`n line number " A_LineNumber)
        Send_modifiers(esc_seq, Modifiers, Lock_list)     ;send the system key with modifiers
        ;modifiers must be sent as key down/up commands for these keys
        Modifiers:=""
        return Modifiers                               ;clear after control sent
      }   ;At this point,esc_seq is lower case, variable length
      ;
      if ((SubStr(esc_seq, 1, 2) = "u+") and (StrLen(esc_seq) = 6))  {   ; look for a Unicode character number
        ;MsgBox,unicode   %esc_seq%    Code_Num= %Code_Num%
        if (SubStr(esc_seq, 3, 6) IsXDigit)   {          ;verify  valid hex number
          Code_Num :=SubStr(esc_seq, 1, 6)
          Code_Num := StrUpper(Code_Num)               ;lower case u+ doesn't cut it....
          Send("{" Code_Num "}")
          Modifiers:=""
          return Modifiers                             ;clear after control sent
        }
        fail_hint:="`n Invalid number format"
        Goto Bad_code
      }
      esc_seq:=Modifiers  esc_seq                   ;add modifies and send whatever is left
      ;MsgBox(" implied press, other=" esc_seq "`n Line number " A_LineNumber
      Send(esc_seq)                                 ;done with this field
      Modifiers:=""
      return Modifiers
    }
    X_click:                                        ; N_Clicks=1, Mouse click or N_Clicks=2, double click
       Mouse_Stop:=1                                ; Stop any movement due to Mougo
        if (esc_seq ="")  {                         ;  click or dblclick with no parameter field, lookup default key
            if (alias_start:=InStr(alias_list, "`nbutdefault="))  { ;find the start of an exact alias match
              alias_start:=InStr(alias_list, "=", , alias_start)+1 ;move up passed the'='
              alias_end:=InStr(alias_list, "`n", , alias_start) ;find the end of match
              esc_seq:= SubStr(alias_list, alias_start, alias_end-alias_start)
            }   ;esc_seq should now have the default button assignment
        }     ;
        if (esc_seq!= "right")  and  (esc_seq!= "left") and !(InStr(Mouse_keys,  esc_seq ".")) { ; recognize right or left as
          fail_hint:="`n Invalid button name"       ; Valid mouse button ?
          Goto Bad_code
        }
        esc_seq:=StrReplace(esc_seq, "button")      ;remove "button from esc_seq. valid parameters are:
        esc_seq:=Modifiers "{Click " esc_seq " " N_Clicks "}" ; modifiers , right,left,r,l,m, x1 and x2
        ;MsgBox(esc_seq )                           ; 'Click does not honor RButton. ??? MButton, Xbutton1, XButton2 ????
        Send esc_seq                                ;send the system key with modifiers
        Modifiers:=""
        return Modifiers       ;and done

    Bad_code:       ;unknown escape sequence
    ; Use a small message box to display escape sequence with error. sending a known defective escape sequence
    Modifiers:=""   ;can create havoc, especially if keys are locked...
    ToolTip("Escape sequence failed.`n" esc_sav fail_hint)
    SetTimer(RemoveToolTip,-5000)                       ;set a one time 5 Second timer to remove tooltip
    SoundPlay "*64", "Wait"                             ;play asterisk sound
    fail_hint:=""                                       ; clear the hint
    return Modifiers
; see https://www.autohotkey.com/boards/viewtopic.php?t=4777   for a tooltip with background color
  }           ;end Parse_esc()
RemoveToolTip()
{
ToolTip()
return
}
;**************************************************************************************************
; Send_modifiers() function is part of the decoder package.
;   1) Prepare system key name to be sent by wrapping the key name in {} and
;   2) If the modifier key is not locked, Prepend{modifier_name down}and append {modifier_name up}
;   3) Send the system key with modifiers.
;   4) return the modified escape sequence
;***************************************************************************************************
Send_modifiers(esc_seq,Mods,Locked_keys)
{
  esc_seq := "{" esc_seq "}"
  if (InStr(Mods, "+") and not InStr(Locked_keys, "shift")) { ;test for shift and not locked
    esc_seq :="{shift down}" esc_seq "{shift up}"
  }
  if (InStr(Mods, "!") and not InStr(Locked_keys, "alt"))  { ;test for alt and not locked
    esc_seq :="{alt down}" esc_seq "{alt up}"
  }
  if (InStr(Mods, "^") and not InStr(Locked_keys, "ctrl")) { ;test for ctrl and not locked
    esc_seq :="{ctrl down}" esc_seq "{ctrl up}"
  }
  if (InStr(Mods, "#") and not InStr(Locked_keys, "lwin")) { ;test for winkey and not locked
    esc_seq :="{lwin down}" esc_seq "{lwin up}"
  }
  SendInput(esc_seq)    ;
  Return esc_seq
}
;#####################################################################################################
;############       End GIDEI Escape sequence Decodeer            ####################################
;#####################################################################################################



