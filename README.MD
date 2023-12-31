Serial Master is an AutoHotkey application that directs Serial port data and/or keyboard data to a designated users application program. The program came into existance due to unresolvable issues that AAC Keys has with some newer serial drivers. AAC keys is/was used by the user community with disabilities to interface Augmentive Assistive devices to the PC world.
Serial Master is a superset of functions found in AAC Keys. There is a PDF file the describes program functions at the level a reasonably sophisticated PC user will understand.
There are 5 component AHK-V2 files to Serial Master:
Serial_Master.ahk is the 'main' component that #includes and initializes other components, edits a task bar menu, launches a Splash Screen, captures keyboard data, hosts a receive loop to invoke a GIDEI Standard decoder and direct input data.
Serial_Master_V2_Splash.ahk Is a timed Splash/static 'About' GUI invoked on startup or via task bar menu.\par
GIDEI_2_Decoder_v2.ahk Is a standardized escape sequence decoder invoked to capture and decode serial input sequences that start with an escape control (0x1B) code. A separate PDF document is available to document decoder function and the new V2 functions. These escape sequences allow virtually full control of the PC via serial port.
Com_Port_Services_V2.ahk This is the Comm port functional component that includes fairly robost error handling as well as a complete Comm Port configuration GUI. A '.CFG file is used to set Comm Port parameters on start-up and may be edited and saved by a GUI invoked in the task bar menu. The GUI also displays live com port control lead state.
Some error conditions also open the Comm GUI with prompts to aid resolution of error conditions that require user intervention. A 'starter' .CFG file is available as part of the Docunentation package.
Serial_Master_Menu_V2.ahk This component is used to select the User application and high level user options on start-up. The module includes a task bar menu invoked GUI that is Profile Oriented, that is, selection of the Users application program sets other GUI options associated with that choice. An additional '.CFG' file is used by the GUI to save edited operating parameters. A 'starter' .CFG file is available as part of the Docunentation package.\par

This application is intended to serve the users of Augmentive Assistve devices as well as Arduino and other hobbiests seeking a simple, self contained PC interface.
}
 
