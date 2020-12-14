' Visual Basic Script Library for use as an include in other Windows Script files.

Option Explicit
'On Error Resume Next

' ************************
'  *** Define Constants ***
'   ************************
' NOTE: const defs continue through header comments...

' #############################################################################
'  ######			DEVELOPER SUMMARY			  ######
'   #############################################################################

' ########################
'  ### Global Variables ###
'   ########################

' ### Strings ###
	Const szLibDesc = "Windows Script Host (VBScript) Library"
	Const szLibVer ="2010.04.15"
	
' ### Objects ###
'	oArgs	- and array of strings containing the command-line parameters that were passed to the script.
'	WshShell- Shell Object; see http://msdn.microsoft.com/library/en-us/script56/html/wsObjWshShell.asp?frame=true
'	WshSysEnv - Shell Environment vars; see http://msdn.microsoft.com/library/en-us/script56/html/wsObjWshEnvironment.asp?frame=true
'	oShellApp - Shell.Application Object; see http://msdn.microsoft.com/library/en-us/shellcc/platform/shell/reference/objects/shell/application.asp?frame=true
'	fs 	- FileSystemObject; see http://msdn.microsoft.com/library/en-us/script56/html/FSOoriFileSystemObject.asp?frame=true
'	oTSLog	= TextStream object for logging script activity

' ############################
'  ### Routine Declarations ###
'   ############################


' ### Script Information ###
'	LibraryVersion() 	- Returns the version of this library (e.g. "2003.12.02")

'	LibraryDesc() 		- Returns the Description of this library, including the szLibDesc value above, the running script name and Library version.

'	Interpreter() 		- Displays the script interpreter (usually WScript.exe).

'	LibPath() 		- Returns the folder path containing the library

'	GetArg(szArgName)	- Returns a command-line argument's value based on its name (format <name><value> is assumed).
'		- Example: if [-file="My file.txt"] is a CL param, calling GetArg("-file=") will return """My file.txt"""; note that only the first match in the list of parameters will be reported

' ### Process Control ###

'	wsRun(szFileName, szWinTitle)
'		- Run the program sFileName and wait for a Window with the title caption szWinTitle to appear
'			Notes: an empty string will not wait for a window; if szWinTitle is not empty and the window does not appear, the script will hang
'	RunBat(szBATFile) - Executes an MS-DOS Batch (.BAT) file and returns the errorlevel from CMD.EXE
'	RunScript(szScript, szParams) - Executes a Windows Script Host compatible script (.JS, .VBS, .WSH, etc.)
'		- Function returns TRUE if the script parser returns zero (no error) and FALSE if the script returns an error.
'	KillProcess(szProcName) - Kills all running processes with the name szProcName (typically the exe name listed in Task Mgr)
'		- Function returns -1 for proc not found, 0 for success or an applicable error code (see constants below)

' ### Window Control ###

' Windows are identified primarily via the captions displayed in the title bar.  
'  As such, if more than one window has the same (or similar) text in the title bar, the wrong application might be controlled.
'  It is therefore of tremendous importance to 1) work with as few windows as possible at one time and 
'  2) include the full window title caption, or as much of it as can be reliably predicted.

'	CloseWindow(szWinTitle)
'	- Select a Window with the title caption szWinTitle and close it (with ALT+F4)

'	WaitOnWindowExit(szWinTitle)
'		- Wait for a Window with the title caption szWinTitle to close; if szWinTitle is not empty and the window does not close, the script will hang.

'	WaitOnWindow(szWinTitle, iSeconds)
'		- Wait for a Window with the title caption szWinTitle to appear and activate it.
'		- Function returns TRUE if the window was found within iSeconds (if iSeconds = 0, script will wait indefinitely)

' Keystroke Generation
'	wsKey(szKeystroke)	- Send szKeystroke to active window
'	wsRepeatKey(szKeystroke, iCount)	- Repeatedly send szKeystroke a total of iCount times.
'	wsKeyString(sKeystrokes)- Send a string as a series of keystrokes
'		- Note that while the wsKey function will support sending multiple keystrokes, problems with timing (some keystrokes are sent too quickly and are discarded) and formatting (some special characters must be enclosed in braces) are resolved in this sub

'  ### System Functions ###

'	SystemName()
'	Shutdown(szShutdownVerb, bWaitforResume) - Performs one of the following functions (requires a copy of WShutdn.exe to be in the script path)
'		- bWaitforResume will wait WShutdn to exit before returning (after resume has completed).
		Const SV_Shutdown = "s"
		Const SV_Restart = "r"
		Const SV_Hibernate = "h"
		Const SV_Standby = "b"
		Const SV_Logoff = "l"

' Operating System Info
'	OSVersion() 		- Provides the base Windows version, such as "Windows 2000" or "Windows XP"
'	FullWindowsVersion() 	- Provides extended OS information, such as "Windows XP Professional Service Pack 1"
'	ServicePackVersion()	- Returns the number representing the installed Windows Service Pack
'	Is2K3()			- Returns TRUE if any variant of Windows Server 2003 or XP x64 is installed
'	IsXP()			- Returns TRUE if any variant of Windows XP is installed
'	Is2K()			- Returns TRUE if any variant of Windows 2000 is installed

' Environment Variables
'	WinDir()		- Returns %WinDir% Environment variable as a string
'	CPU()			- Returns %PROCESSOR_ARCHITECTURE% Environment variable as a string

' Modem Control
'	OpenModemProperties(szModemName, bOpen) - Opens / closes the modem control panel applet and gets properties for a modem with the friendly name szModemName
'		- Function returns TRUE if the configuration window was successfully displayed.
'		- If called with bOpen = False, function will close Properties window and applet

'  ### Registry Access ###

'	ScriptRegRead(szValueName) - Retrieves a value stored in a script-specific registry key
'	ScriptRegWrite(szValueName, Value) -  Writes the supplied Value to an entry named szValueName to a script-specific registry key 
'	RegValueExists(FullRegKeyValuePath) - Returns TRUE if the specified value exists.
'	RegValueGet(FullRegKeyValuePath)
'	AddReg(szRegFile) - Merges the contents of a .REG file into the host system registry using Regedit.exe

'  ### Low-Level File Operations ###

'	GetFileAttributes(szFileName, iAttribute) - Uses oShellApp object to report file and folder attributes
'		- See Extended Properties Constant list below.
'	Attrib(szFileName, iAttrib, bValue)
'		- Writes to a shorter list of Attributes than GetFileAttributes (see constant list below)
'	OpenLogFile(szScriptName) - Opens a log file using the szScriptName paramter as a guide (file extension will be changed to ".log")
'	WriteLog(szLogMsg) - Writes a string and timestamp to an open log file.
'	OpenTxtFile(szFileName, IOMode) - Returns a TextStream Object for reading from the specified filename. see IOMode tristate constants below
'		- Usage: Set MyTextStream = OpenTxtFile("C:\Boot.ini", ForReading)
'		- For information on using TextStream objects, see http://msdn.microsoft.com/library/en-us/script56/html/jsobjtextstream.asp?frame=true
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
'	ExtractFilePath(szFullFileName) - returns the path to the folder containing the file or folder specified
'	ExtractFileName(szFullFileName) - returns only the name of the file specified with path information removed.
'	ExtractFileExt(szFullFileName, bBackward) - returns only the MS-DOS extension of the file specified with path information removed.
'		- If Called with bBackward=TRUE, function returns the part of the file name with no extension
'	CheckPath(szFileName, bAddQuotes) - verifies that a filename contains a full path; the script local path will be substituted if no absolute or relative path is specified.
'		- if bAddQuotes is TRUE, the returned string will be enclosed in double quotes.
'	ReplaceFile (szSource, szTarget) ' Target is file is over-written by source file, which is then deleted.


'  ### Dialog Boxes ###

'	GetParam(szMesg, szDefault) - Displays an Input Box (with the prompt szMessage) and returns the string entered by the user
'		- The string szDefault will be displayed as the default value in the Input Box
'	Confirm(szPrompt) - Displays an Message Box (with a prompt asking the user to confirm the action szPrompt)
'		- The function returns TRUE if the Yes button is clicked and FALSE for all other results.
'	CheckErrWarn(szOperation, szSourceRoutine)
'		- Reports script run-time errors; NOTE: On Error Resume Next is required
'	CriticalError(szErrMsg)  - Displays an Message Box (with a prompt indicating that the error szErrMsg has occurred)
'		- Note: script execution will terminate when CriticalError is called and the script will return the value ERR_GENERAL_FAILURE
'	About() - Displays a dialog describing the script function and usage.
' 		- Note: you should write to these vars *before* calling about
		Dim szScrTitle		'Friendly name of the script
		Dim szScrVersion	'Script version
		Dim szScrAuthor		'Sctipr author
		Dim szScrDescr		'Verbose description of what the script does
		Dim szScrComment	'Notes to the user
		Dim szScrUsage		'Command-line parameters supported.
'	CheckAbout() - Automatically displays About UI when -h, -H or -? are detected as command-line parameters.  Note that this will prevent execution of a script used with this library when the help command line arguments are included.


'  ### Cryptographic Functions ###

	Const c_CryptKeyPath="HKLM\SOFTWARE\Microsoft\Windows Scripting Host\Library.vbs.CryptKey"
'	szEncrypt(szClearText) - Encrypt szClearText and return the CypherText; XOR encryption using a static key (above)
'	GetBit(charToCheck, iBitToGet) - Retrieve bit values for a given character 
'	InitKey(szKey) - reset the cryptographic key

'  ### Network Functions ###
'	ActiveNetwork - returns a boolean value indicating whether a network adapter is present with an active link state.
'	Ping(szHost) - returns a boolean value indicating whether TCP echo response was received
'	NetSend(szHost, szMessage) - displays a dialog to all users on the remote system szHost with message szMessage (requires NT/2K/XP/2K3 with Messenger service active) 

'  ### Misc Routines ###

' 	ReplaceSubstr(szSource, szFind, szReplace) - replace all occurrences of szFind (found in szSource) with szReplace.
'	DbgOut(szMsg) 	' Print debug messages in when Debug Mode is active (i.e., when bDebug is TRUE).
'	CheckBit(iBitMask, iBitNum)	'Returns TRUE if the bit iBitNum in the value iBitMask is set.
'	EmptyArray(arrTest) - Returns TRUE is the dynamic array arrTest is uninitialized

' #############################################################################


' *** Script Exit Codes ***
Const ERR_OK              = 0
Const ERR_GENERAL_FAILURE = 1

' *** Run-time Consts ***
Const cDefaultDebug = False
Dim bDebug: bDebug = cDefaultDebug

' *** XML Entities / Codes ***
'  used with RExec, among other things
Const cQuotes = "&quot;"
Const cPercent = "&#37;"
Const cAmpersand = "&amp;"

' *** Message Box Codes ***
' Base Types
Const MB_OKOnly           = 0  
Const MB_OKCancel         = 1  
Const MB_AbortRetryIgnore = 2   
Const MB_YesNoCancel      = 3  
Const MB_YesNo            = 4  
Const MB_RetryCancel      = 5
' Option Flags (these are additive)  
Const MB_Critical         = 16
Const MB_Question         = 32
Const MB_Exclamation      = 48
Const MB_Information      = 64
Const MO_SystemModal      = 4096
Const MO_ForeGround       = 65536
' Return Codes
Const MR_OK               = 1 
Const MR_Cancel           = 2  
Const MR_Abort            = 3 
Const MR_Retry            = 4 
Const MR_Ignore           = 5  
Const MR_Yes              = 6 
Const MR_No               = 7 

'Basic File Properties (used With Attrib()
' See http://msdn.microsoft.com/library/en-us/script56/html/jsproattributes.asp?frame=true
'Const P_Normal 	= 0 'Normal file. No attributes are set. 
Const P_ReadOnly	= 1 'Read-only file. Attribute is read/write. 
Const P_Hidden 		= 2 'Hidden file. Attribute is read/write. 
Const P_System 		= 4 'System file. Attribute is read/write. 
'Const P_Volume 	= 8 'Disk drive volume label. Attribute is read-only. 
'Const P_Directory 	= 16 ' Folder or directory. Attribute is read-only. 
Const P_Archive 	= 32 'File has changed since last backup. Attribute is read/write. 
'Const P_Alias 		= 1024 'Link or shortcut. Attribute is read-only. 
'Const P_Compressed 	= 2048 'Compressed file. Attribute is read-only. 

' Extended Properties Constants from Shell Folder Object
'  see http://msdn.microsoft.com/library/en-us/shellcc/platform/shell/reference/objects/shellfolderitem/extendedproperty.asp?frame=true
Const EP_Name  = 0
Const EP_Size  = 1
Const EP_Type  = 2
Const EP_Date_Modified  = 3
Const EP_Date_Created  = 4
Const EP_Date_Accessed  = 5
Const EP_Attributes  = 6
Const EP_Status  = 7
Const EP_Owner  = 8
Const EP_Author  = 9
Const EP_Title  = 10
Const EP_Subject  = 11
Const EP_Category  = 12
Const EP_Pages  = 13
Const EP_Comments  = 14
Const EP_Copyright  = 15
Const EP_Artist  = 16
Const EP_Year     = 18
Const EP_Track_Number  = 19
Const EP_Genre    = 20
Const EP_Duration  = 21
Const EP_Bit_Rate  = 22
Const EP_Protected  = 23
Const EP_Camera_Model  = 24
Const EP_Date_Picture_Taken  = 25
Const EP_Dimensions  = 26
'Const EP_Not_used_1  = 27
'Const EP_Not_used_2 = 28
'Const EP_Not_used_3 = 29
Const EP_Company  = 30
Const EP_Description   = 31
Const EP_File_Version  = 32
Const EP_Product_Name  = 33
Const EP_Product_Version  = 34

'Process Termination Constants
Const cTERM_SUCCESS 		= 0  'Successful completion 
Const cTERM_ACCESSDENIED 	= 2  'Access denied 
Const cTERM_INSUFFPRIV 		= 3  'Insufficient privilege 
Const cTERM_UNKNOWN 		= 8  'Unknown failure 
Const cTERM_PATHNOTFOUND 	= 9  'Path not found 
Const cTERM_INVALID 		= 21 'Invalid parameter 

' *** Registry Access Constants ***
' Note: The registry provider for Windows Server 2003 is hosted in LocalService, not the LocalSystem. Therefore, using the Windows Server 2003 family you cannot obtain information remotely from HKEY_CURRENT_USER
Const cHiveHKCU 	= &H80000001	'HKEY_CURRENT_USER
Const cHiveHKCR 	= &H80000000	'HKEY_CLASSES_ROOT
Const cHiveHKLM 	= &H80000002	'HKEY_LOCAL_MACHINE
Const cHiveHKU 		= &H80000003	'HKEY_USERS
Const cHiveHKCC		= &H80000005	'HKEY_CURRENT_CONFIG
Const cHiveHKDD		= &H80000006	'HKEY_DYN_DATA


' *** Init Objects ***
' For information on CreateObject method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsMthCreateObject.asp?frame=true
Dim fs: Set fs = CreateObject("Scripting.FileSystemObject")
Dim WshShell: Set WshShell = WScript.CreateObject("WScript.Shell")
Dim WshSysEnv: Set WshSysEnv = WshShell.Environment("PROCESS")
'Set WshSysEnv = WshShell.Environment("SYSTEM")
Dim oShellApp: Set oShellApp = CreateObject("Shell.Application")
Dim oArgs: Set oArgs = WScript.Arguments
Dim oTSLog
WriteLog WScript.ScriptName & " Loaded Library: " & LibraryDesc

' **************************
'  *** Script Information ***
'   **************************

Function LibraryVersion()
	LibraryVersion = szLibVer
End Function

Function LibraryDesc()
	LibraryDesc = szLibDesc + " (Library.vbs, version " + szLibVer + ")"
End Function

Function Interpreter()
	Interpreter = Right(WScript.FullName, len(WScript.FullName) - len(WScript.Path) - 1)
End Function

Function LibPath() ' Returns the folder path containing the library
	Dim szTemp: szTemp = WScript.ScriptFullName
	szTemp = Mid(szTemp, 1, InStr(szTemp, WScript.ScriptName) - 1)
	LibPath = szTemp
End Function

Function GetArg(szArgName)
	Dim arg
	GetArg = ""
	szArgName = UCase(szArgName)
	For Each arg in oArgs
		If Left(UCase(arg), Len(szArgName)) = szArgName Then
			If UCase(arg) = szArgName Then Exit Function
			GetArg = Mid(arg, Len(szArgName) + 1)
		End If
	Next
End Function


' ***********************
'  *** Process Control ***
'   ***********************

Function wsRun(szFileName, szWinTitle)
	WshShell.Run szFileName
	' For documentation on the Run method, see http://msdn.microsoft.com/library/en-us/script56/html/wsmthrun.asp?frame=true
	If szWinTitle <> "" Then
		wsRun = WaitOnWindow(szWinTitle, 0) ' wait for the target Window title to appear
	Else
		wsRun = True
	End If
	WScript.Sleep 500
End Function

Function RunBat(szBATFile)
	RunBat = WshShell.Run("%SYSTEMROOT%\system32\cmd.exe /c " & CheckPath(szBATFile, True),,TRUE)
End Function

Function RunScript(szScript, szParams)
	DbgOut "Running Script: " & szScript & " [" & szParams & "]"
	RunScript = (WshShell.Run("%SYSTEMROOT%\system32\WScript.exe " & CheckPath(szScript, True) & " " + szParams,,TRUE) = 0)
End Function

Function KillProcess(szProcName)
	' for more information on Win32_Process, see http://msdn.microsoft.com/library/en-us/wmisdk/wmi/win32_process.asp?frame=true
	Dim Process
	KillProcess = (-1) ' process not found is the default
	
	For Each Process in GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_process")
		If UCASE(Process.Name) = ucase(szProcName) Then _
			KillProcess = Process.Terminate()
	Next 
End Function


' **********************
'  *** Window Control ***
'   **********************

Sub CloseWindow(szWinTitle)
	' Actively close a Window
	If szWinTitle = "" Then Exit Sub
	WaitOnWindow szWinTitle, 0 ' wait for the target Window title to appear
	wsKey "%{F4}"
End Sub

' For information on AppActivate method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsmthappactivate.asp?frame=true

Sub WaitOnWindowExit(szWinTitle)
	'Wait until the application has exited - Check every half second
	If szWinTitle = "" Then Exit Sub
	WshShell.AppActivate szWinTitle
	While WshShell.AppActivate(szWinTitle) = TRUE
		WshShell.AppActivate szWinTitle
		WScript.Sleep 500
	Wend
End Sub

Function WaitOnWindow(szWinTitle, iSeconds)
	'Wait until the application has loaded - Check every quarter second
	If szWinTitle = "" Then Exit function
	const WaitInterval = 250 ' milliseconds between each check for the target window title
	WaitOnWindow = FALSE
	If iSeconds = 0 then
		While (WshShell.AppActivate(szWinTitle) = FALSE)
			WScript.Sleep WaitInterval 
		Wend
		WaitOnWindow = TRUE
	Else
		Dim i: i = 0
		While i < ((1000/WaitInterval) * iSeconds)
			If WshShell.AppActivate(szWinTitle) Then
				WshShell.AppActivate szWinTitle
				WaitOnWindow = TRUE
				Exit Function
			Else
				i = i + 1
				WScript.Sleep WaitInterval 
			End If
		Wend
	End If
End Function

' *** Keystroke Generation ***
' For information on SendKeys method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsMthSendKeys.asp?frame=true

Sub wsKey(szKeystroke)
	WshShell.SendKeys szKeystroke
	WScript.Sleep 75
End Sub

Sub wsRepeatKey(szKeystroke, iCount)
	If iCount < 1 Then Exit Sub
	Dim i
	For i = 1 To iCount
		wsKey(szKeystroke)
	Next
End Sub

Sub wsKeyString(sKeystrokes)
	Dim szTempChar
	'DbgOut "Sending " + sKeystrokes + "..." 

	Dim i
	For i = 1 To len(sKeystrokes)
		WScript.Sleep 5

		szTempChar = Mid(sKeystrokes,i,1)
		Select Case szTempChar
			Case vbTab
				szTempChar = "{TAB}"
			Case "("
				szTempChar = "{(}"
			Case ")"
				szTempChar = "{)}"
			Case "{"
				szTempChar = "{{}"
			Case "}"
				szTempChar = "{}}"
			Case "["
				szTempChar = "{[}"
			Case "]"
				szTempChar = "{]}"
			Case "+"
				szTempChar = "{+}"
			Case "^"
				szTempChar = "{^}"
			Case "%"
				szTempChar = "{%}"
			Case "~"
				szTempChar = "{~}"
			Case vbCr
				szTempChar = "{ENTER}"
		End Select

		WshShell.SendKeys szTempChar
	Next
	WScript.Sleep 75
End Sub


' ************************
'  *** System Functions ***
'   ************************

Function SystemName()
	Dim Network ' For information on Network Object, see http://msdn.microsoft.com/library/en-us/script56/html/wsObjWshNetwork.asp?frame=true
	Set Network = WScript.CreateObject("WScript.Network")
	SystemName=Network.ComputerName
End Function

Sub Shutdown(szShutdownVerb, bWaitforResume)
	szShutdownVerb = LCase(szShutdownVerb)
	dim szWorkingPath: szWorkingPath = "d:\work\WShutdn.exe"
	Select Case szShutdownVerb
		Case "r", "s", "l", "h", "b"
			If Not fs.FileExists(szWorkingPath) Then _
				szWorkingPath = LibPath & "WShutdn.exe"
			
			WshShell.Run """" & szWorkingPath & """ " & szShutdownVerb,,bWaitforResume
		Case Else
			DbgOut "Error in Shutdown: verb """ & szShutdownVerb & """ is not supported"
	End Select
End Sub

' *** Operating System Info ***

Function OSVersion()
	Dim OSVersionRegKey
	OSVersionRegKey=RegValueGet("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentVersion")
	Select Case OSVersionRegKey
		Case "4.0"
			OSVersion="Windows NT 4.0"
		Case "5.0"
			OSVersion="Windows 2000"
		Case "5.1"
			OSVersion="Windows XP"
		Case "5.2"
			OSVersion="Windows Server 2003"
		Case "6.0"
			OSVersion="Windows Vista"
		Case Else
			OSVersion="Other known as """ & OSVersionRegKey & """"
	End Select
End Function

Function FullWindowsVersion()
	Dim WindowsVersion,ProductVersion,TerminalServerMode
	WindowsVersion=OSVersion
	ProductVersion=RegValueGet("HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType")
	TerminalServerMode=RegValueGet("HKLM\System\CurrentControlSet\Control\Terminal Server\TSAppCompat")


	Select Case WindowsVersion
		Case "Windows NT 4.0"
			Select Case ProductVersion
				Case "ServerNT"
					If (TerminalServerMode=1) Then
						ProductVersion=" Terminal Server"
					Else
						ProductVersion=" Server"
					End If
				Case "LanmanNT"
					ProductVersion=" Server (Domain Controller)"
				Case "WinNT"
					ProductVersion=" Workstation"
				Case Else
					'ProductVersion=""
			End Select
		Case "Windows 2000"
			Select Case ProductVersion
				Case "ServerNT"				
					If (TerminalServerMode=1) Then
						ProductVersion=" Server (Terminal Services Application Server Mode)"
					Else
						ProductVersion=" Server"
					End If
				Case "WinNT"
					ProductVersion=" Professional"
				Case Else
					'ProductVersion=""
			End Select
		Case "Windows XP"
			Select Case ProductVersion
				Case "ServerNT"
					ProductVersion=" Server"
				Case "WinNT"
					ProductVersion=" Professional"
				Case Else
					'ProductVersion=""
			End Select
		Case "Windows Server 2003"
			Select Case ProductVersion
				Case "ServerNT"
					If (TerminalServerMode=1) Then
						ProductVersion=" Server (Terminal Services Application Server Mode)"
					Else
						ProductVersion=" Server"
					End If
				Case "WinNT"
					ProductVersion=" Professional"
				Case Else
					'ProductVersion=""
			End Select 

		Case "Windows Longhorn"
			Select Case ProductVersion
				Case "ServerNT"
					ProductVersion=" Server"
				Case "WinNT"
					ProductVersion=" Professional"
				Case Else
					'ProductVersion=""
			End Select
		Case Else
			'WindowsVersion=ProductVersion=""
	End Select
	Dim ServicePackVer
	ServicePackVer=ServicePackVersion()
	If (ServicePackVer>0) Then
		ServicePackVer=" Service Pack " & ServicePackVer
	Else
		ServicePackVer=""
	End If

'	WScript.echo "ServicePackVer=" & ServicePackVer

	FullWindowsVersion=WindowsVersion & ProductVersion & ServicePackVer
End Function

Function ServicePackVersion()
	Dim CSDVersion,oRE,oMatches,oMatch,i,SPVer

	Set oRE = New RegExp
	oRE.Global=True
	oRE.IgnoreCase=True

	If (RegValueExists("HKLM\System\CurrentControlSet\Control\Windows\CSDVersion")=0) Then
		CSDVersion=0
	Else
		CSDVersion=RegValueGet("HKLM\System\CurrentControlSet\Control\Windows\CSDVersion")
		If (CSDVersion="") Then
			CSDVersion=0
		ElseIf (CSDVersion>0) Then
			CSDVersion=Hex(CSDVersion)
		End If
	End If

	SPVer=0

	oRE.Pattern="(\d)00$"
	Set oMatches=oRE.Execute(CSDVersion)
	i=0
	For Each oMatch In oMatches
		SPVer=left(CSDVersion,1)
		i=i+1
	Next

	If (i>0) Then
		ServicePackVersion=SPVer
	Else
		ServicePackVersion=0
	End If
End Function

Function Is2K3()
	Is2K3 = (OSVersion = "Windows Server 2003")
End Function

Function IsXP()
	IsXP = (OSVersion = "Windows XP")
End Function

Function Is2K()
	Is2K = (OSVersion = "Windows 2000")
End Function

' *** Environment Variables ***

Function WinDir
	If IsXP Or Is2K Or Is2K3 Then
		WinDir = WshSysEnv("SystemRoot")
	Else
		WinDir = WshSysEnv("WinDir")
	End If
End Function

Function CPU
	Dim objWMIService: Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colItems: Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
	Dim objItem, iArch
	For Each objItem in colItems
		iArch = objItem.Architecture
	Next
	Set colItems = Nothing
	Set objWMIService = Nothing
	Select Case iArch
		Case 0: CPU = "X86"
		Case 1: CPU = "MIPS"
		Case 2: CPU = "ALPHA"
		Case 3: CPU = "POWERPC"
		Case 6: CPU = "IA64"
		Case 9: CPU = "X64"
		Case Else: CPU = "UNKNOWN"
	End Select

	' This does not work when running in 32-bit script host
	'If IsXP Or Is2K Or Is2K3 Then
	'	CPU = UCASE(Cstr(WshSysEnv("Processor_Architecture")))
	'Else
	'	CPU = "X86"
	'End If
End Function

' *** Modem Control ***

Function OpenModemProperties(szModemName, bOpen)
	Dim szWinTitle

	Select Case OSVersion
		Case "Windows XP", "Windows 2000", "Windows Server 2003"
			szWinTitle = "Phone and Modem Options"
		Case "Windows Longhorn"
		Case Else
			CheckErrWarn "Unsupported OS (" & OSVersion & ")", "OpenModemProperties"
			Exit Function
	End Select
	If bOpen Then
		If CPU = "AMD64" Then
			wsRun "%SystemRoot%\SysWOW64\Control.exe telephon.cpl", szWinTitle
		Else
			wsRun "%SystemRoot%\System32\Control.exe telephon.cpl", szWinTitle
		End If
		wsKey "^{TAB}"
		wsKeyString szModemName
		wsKey "%p"
		OpenModemProperties = WaitOnWindow(szModemName, 20)
	Else
		WaitOnWindow szModemName, 0
		wsKey "{ENTER}"
		CloseWindow szWinTitle
		OpenModemProperties = TRUE
	End If
End Function


' ***********************
'  *** Registry Access ***
'   ***********************
' For info on the WMI StdRegProv Class, see http://msdn.microsoft.com/library/en-us/wmisdk/wmi/stdregprov.asp?frame=true

Function ScriptRegRead(szValueName)
	Dim szReadValue: szReadValue = ""
	Dim szScriptKeypath: szScriptKeypath = "HKCU\Software\Microsoft\Windows Script\" & WScript.ScriptName
	szReadValue = RegValueGet(szScriptKeypath & "\" & szValueName)
	DbgOut "Read from Script Key value [" & szValueName &"]: " & szReadValue
	ScriptRegRead = szReadValue
End Function

Function ScriptRegWrite(szValueName, Value)
	Dim bSuccess: bSuccess = false
	Dim szScriptKeypath: szScriptKeypath = "HKCU\Software\Microsoft\Windows Script\" & WScript.ScriptName
	'If (RegValueExists(FullRegKeyValuePath)>0) Then
	'	WshShell.RegWrite(FullRegKeyValuePath, szValueName)
	'End If
	WshShell.RegWrite szScriptKeypath & "\" & szValueName, Value
	bSuccess = (RegValueGet(szScriptKeypath & "\" & szValueName) = Value)
	ScriptRegWrite = bSuccess
End Function

Function RegValueExists(FullRegKeyValuePath)
' Note: The registry provider for Windows Server 2003 is hosted in LocalService, not the LocalSystem. Therefore, using the Windows Server 2003 family you cannot obtain information remotely from HKEY_CURRENT_USER
	On Error Resume Next
	Dim oReg,Hive,KeyPath,RegValue,ValueNames(),ValueTypes(),oRE,Enumeration,Counter,Arch
	Set oReg=GetObject("Winmgmts:root\default:StdRegProv")
	
	Set oRE = New RegExp
	oRE.IgnoreCase=True
	oRE.Global=True
	oRE.Pattern="^(\w+)\\(.*)\\(.+)$"

	Hive=oRE.Replace(FullRegKeyValuePath,"$1")
	KeyPath=oRE.Replace(FullRegKeyValuePath,"$2")
	RegValue=oRE.Replace(FullRegKeyValuePath,"$3")

	DbgOut "Hive=" & Hive
	DbgOut "KeyPath=" & KeyPath
	DbgOut "RegValue=" & RegValue

	Select Case Hive
		Case "HKCR"
			Hive = cHiveHKCR
		Case "HKCU" ' 
			Hive = cHiveHKCU
		Case "HKLM"
			Hive = cHiveHKLM
		Case "HKU"
			Hive = cHiveHKU
		Case "HKCC"
			Hive = cHiveHKCC
		Case "HKDD"
			Hive = cHiveHKDD
		Case Else
			Exit Function
	End Select
	
	RegValueExists = 0

	Enumeration=oReg.EnumValues(Hive, KeyPath, ValueNames, ValueTypes)

 	If ((Enumeration=0) AND (Err.Number=0) AND (Not EmptyArray(ValueNames)) AND (Not IsNull(ValueTypes))) Then
		For Counter = 0 To UBound(ValueNames)
			DbgOut "Counter=" & CStr(Counter)
			If (LCase(ValueNames(Counter)) = LCase(RegValue)) Then
				DbgOut "LCase(ValueNames(" & Counter & "))=" & LCase(ValueNames(Counter))
				DbgOut "LCase(RegValue)=" & LCase(RegValue)
				DbgOut "Setting RegValueExists to 1"
				RegValueExists = 1
				Exit For
			End If
		Next
	End If
	
	CheckErrWarn "Enumerating values in Key: " & KeyPath, "RegValueExists"
End Function

' For Information on the RegDelete, RegRead, and RegWrite methods in the WshShell object,
'   see http://msdn2.microsoft.com/en-us/library/aew9yb99.aspx 

Function RegValueGet(FullRegKeyValuePath)
	On Error Resume Next

	Dim RegKeyValue

	If (RegValueExists(FullRegKeyValuePath)>0) Then
		RegKeyValue=WshShell.RegRead(FullRegKeyValuePath)

		If CheckErrWarn("Reading registry value: " & FullRegKeyValuePath, "RegValueGet") Then
			RegValueGet=0
		Else
			RegValueGet=RegKeyValue
		End If
	Else
		RegValueGet=0
	End If
End Function

Function RegValueDel(FullRegKeyValuePath)
	On Error Resume Next

	RegValueDel=false
	If (RegValueExists(FullRegKeyValuePath)>0) Then

		WshShell.RegDelete FullRegKeyValuePath

		If (RegValueExists(FullRegKeyValuePath)>0) Then
			If CheckErrWarn("Deleting registry value: " & FullRegKeyValuePath, "RegValueDel") Then
			Else
				RegValueGet=true
			End If
		End If 
	Else
		RegValueDel=true
	End If
End Function

Sub AddReg(szRegFile)
	WshShell.Run "%SYSTEMROOT%\regedit.exe -s " & CheckPath(szRegFile,TRUE),,TRUE
End Sub


' *********************************
'  *** Low-Level File Operations ***
'   *********************************

Function GetFileAttributes(szFile, iAttribute)
	DbgOut "Getting Attributes for path: " & szFile
	szFile = CheckPath(szFile, False)

	DbgOut "Setting oShFldr to " & ExtractFilePath(szFile)
	Dim oShFldr: Set oShFldr = oShellApp.Namespace(ExtractFilePath(szFile))
	Dim szFileName, i

	szFile = ExtractFileName(szFile)
	For Each szFileName in oShFldr.Items
		If UCASE(szFileName) = UCASE(szFile) Then
			GetFileAttributes = oShFldr.GetDetailsOf(szFileName, iAttribute)
			DbgOut "iAtrtribute: " & CStr(iAttribute) & " = " & CStr(GetFileAttributes)
			Exit Function
		End If
	Next
	GetFileAttributes = (-1) ' File name was not matched.
End Function

Sub Attrib(szFileName, iAttrib, bValue)
	szFileName = Checkpath(szFileName,False)
	If fs.FileExists(szFileName) Then
		Dim f: Set f = fs.GetFile(szFileName)
		If (f.attributes And iAttrib) <> bValue Then
			If Not bValue Then iAttrib = iAttrib * (-1)
			f.attributes = f.attributes + iAttrib
		End If
		Set f = Nothing
	End If
End Sub

Sub OpenLogFile(szScriptName)
	Set oTSLog = OpenTxtFile( _
		LibPath & ExtractFileExt(szScriptName, True) & ".log", _
		ForAppending)
	CheckErrWarn "Opening text stream: " & LibPath & ExtractFileExt(szScriptName, True) & ".log", "OpenLogFile"
End Sub

Sub WriteLog(szLogMsg)
	If IsObject(oTSLog) Then _
		oTSLog.WriteLine CStr(Now) & vbTab & szLogMsg
	If lcase(Interpreter) = "cscript.exe" Then _
		WScript.Echo szLogMsg
End Sub

Function OpenTxtFile(szFileName, IOMode)
    	szFileName = CheckPath(szFileName, FALSE)
	
	' Skip debug output to file if we're opening the script log (otherwise script will stick in infinite loop)
	If IsObject(oTSLog) Then _
		DbgOut "Opening file [" + szFileName + "]"

	Const UnicodeUseDefault = -2, UnicodeTrue = -1, UnicodeFalse = 0
	If Not fs.FileExists(szFileName) Then
        	Select Case IOMode
	    	    	Case ForReading  ' Complain
	                	CriticalError "Source File: [" & szFileName & "] does not exist."
	    	    	Case ForWriting, ForAppending
	                	fs.CreateTextFile szFileName       'Create a file
    		End Select
    	ElseIf IOMode = ForWriting Then ' Delete file that will be over-written
        	' Do Nothing
    	End If
    
    	Dim f: Set f = fs.GetFile(szFileName)
    	Set OpenTxtFile = f.OpenAsTextStream(IOMode, UnicodeUseDefault)
End Function

Function ExtractFilePath(szFullFileName)
	Dim szTemp: szTemp = szFullFileName
	While Instr(szTemp, "\") > 0
		' So long as there are still backslashes in the string...
		ExtractFilePath = ExtractFilePath & Left(szTemp, Instr(szTemp, "\"))
		' Append the next section of the file path to the output string...
		szTemp = Right(szTemp, _
			Len(szTemp) - Instr(szTemp, "\"))
		' ...and trim off leading text up to and including the backslash.
	Wend
End Function

Function ExtractFileName(szFullFileName)
	ExtractFileName = szFullFileName
	While Instr(ExtractFileName, "\") > 0
		' So long as there are still backslashes in the string...
		ExtractFileName = Right(ExtractFileName, _
			Len(ExtractFileName) - Instr(ExtractFileName, "\"))
		' trim off leading text up to and including the backslash.
	Wend
End Function

Function ExtractFileExt(szFullFileName, bBackward)
	ExtractFileExt = ExtractFileName(szFullFileName) ' just in case name was passed in with path info.
	dim idx: idx = InStr(szFullFileName, ".")
	If idx = 0 Then
		If Not bBackward Then _
			ExtractFileExt = "" ' no extension to report
		Exit Function
	Else
		If bBackward Then
			ExtractFileExt = Mid(szFullFileName, 1, idx -1)
		Else
			ExtractFileExt = Mid(szFullFileName, idx)
		End If
	End If
	
End Function

Function CheckPath(szFileName, bAddQuotes)
	If InStr(szFileName, "\") = 0 Then 
		szFileName = LibPath + szFileName
	Else
		' Do nothing
	End If
	If (InStr(szFileName, """") <> 1) And bAddQuotes Then 
		CheckPath = """" + szFileName + """"
	Else
		CheckPath = szFileName
	End If
End Function

Sub ReplaceFile (szSource, szTarget) ' Target is over-written by source
    fs.MoveFile szTarget, szTarget & ".del" 
    fs.MoveFile szSource, szTarget
    fs.DeleteFile szTarget & ".del"   ' Original file can be preserved by commenting this line
End Sub


' ********************
'  *** Dialog Boxes ***
'   ********************

' For information on Popup method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsmthpopup.asp?frame=true

Function GetParam(szMesg, szDefault)
    GetParam = InputBox("Please provide the following parameter:" & vbNewLine & "  " & szMesg, "Input needed [" & WScript.ScriptName & "]", szDefault)
End Function

Function Confirm(szPrompt)
    Confirm = (MsgBox("Are you sure you want to "& szPrompt & "?", _
                     MB_Question + MB_YesNo, WScript.ScriptName) = MR_Yes)
End Function

Function CheckErrWarn(szOperation, szSourceRoutine)
' Function Returns True if an error is encountered and reports the error to the operator.
	CheckErrWarn = False
	
	IF Err.Number<>0 Then
		Dim szErrDetails: szErrDetails = "Error #" & Err & ": " & Err.Description
		CheckErrWarn = True
		If IsObject(oTSLog) Then 
			WriteLog "***** Script Error *****"
			WriteLog "Routine: " & szSourceRoutine & "; Operation: " & szOperation
			WriteLog "Details: " & szErrDetails
			WriteLog "***** Resuming Script Execution *****"
		End If
		Warning "in routine """ & szSourceRoutine & _
			""" while attempting the operation:" & _ 
			vbNewLine & vbTab & szOperation & _
			vbNewLine & vbNewLine & vbTab & szErrDetails
		Err.Clear
	End If
End Function

Sub Warning(szWrnMsg)
	If MsgBox("An error was encountered: " & szWrnMsg & _
		vbNewLine & vbNewLine & "Click OK to continue script Execution.", _
		MB_Exclamation + MB_OkCancel, WScript.ScriptName + ": ERROR") = MR_Cancel Then _
			CriticalError "User aborted script execution following run-time error!"
End Sub

Sub CriticalError(szErrMsg)
	If IsObject(oTSLog) Then
		WriteLog "CRITICAL ERROR: " & szErrMsg
		WriteLog "***** Script Terminated *****"
		oTSLog.Close
	End If
	'If bDebug Then _
	'	Wshshell.Run LibPath & ExtractFileExt(WScript.ScriptName, True) & ".log"
	MsgBox szErrMsg & vbNewLine & "Terminating script.", MB_Critical, WScript.ScriptName + ": ERROR"
	WScript.Quit ERR_GENERAL_FAILURE
' For information on Quit method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsMthQuit.asp?frame=true
End Sub

Sub About()
	Dim szMessage: szMessage = szScrTitle & " [" & WScript.ScriptName &  "]" & vbNewLine
	If szScrVersion <> "" Then _
		szMessage = szMessage & "Version: " & szScrVersion
	If szScrAuthor <> "" Then _
		szMessage = szMessage & "   Author: " & szScrAuthor
	szMessage = szMessage & vbNewline & _
			"  using " & szLibDesc & " version " & szLibVer & vbNewLine
	If szScrDescr <> "" Then _
		szMessage = szMessage & vbNewline & szScrDescr & vbNewline
	If szScrUsage <> "" Then _
		szMessage = szMessage & vbNewline & "Usage:" & vbNewLine & szScrUsage & vbNewline
	If szScrComment <> "" Then _
		szMessage = szMessage & vbNewline & "Note: " & szScrComment & vbNewline
	MsgBox szMessage, MB_Information, WScript.ScriptName + ": About"
' For information on Quit method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsMthQuit.asp?frame=true
End Sub

Sub CheckAbout()

	Dim arg, allargs
	allargs = ""
	WriteLog WScript.ScriptName & " Intercepting help on command-line."
	For Each arg in oArgs
		allargs = allargs & " " & arg
		If (arg = "-?") or (UCase(arg) = "-H") Then 
			About
			WScript.Quit ERR_OK
		End If
	Next
	DbgOut "Core Script: " & WScript.ScriptName & " v" & szScrVersion & " - Run with params: [" & allargs & "]"
End Sub


Sub Beep()
	Dim strSoundFile: strSoundFile = WinDir & "\Media\ding.wav"
	Dim strCommand: strCommand = "sndrec32 /play /embedding /close " & chr(34) & strSoundFile & chr(34)
	wshShell.Run strCommand, 0, True
End Sub

' *******************************
'  *** Cryptographic Functions ***
'   *******************************

Function szEncrypt(szClearText)
	Dim i, iKeyLen
	Dim iCypherChar, iPwr
	Dim bKeyBit, bClearBit
	Dim szKey, charKey, charClear

	' simple bit-wise XOR encryption against a static key
	' XOR		XOR
	' M K C		C K M
	'--------------------
	' 0 0 0		0 0 0
	' 0 1 1		1 1 0
	' 1 0 1		1 0 1
	' 1 1 0		0 1 1

	szKey = CStr(RegValueGet(c_CryptKeyPath))
	dbgout "szKey is " & szKey
	If szKey = "0" Then _
		szKey = InitKey(szKey)
	iKeyLen = Len(szKey)
	i = 0
	While i < Len(szClearText)
		i = i + 1
		charKey = Mid(szKey, (i Mod iKeyLen), 1)
		charClear = Mid(szClearText, i, 1)

		If charKey = charClear Then ' encryption would generate a null; that's bad; need a new key.
			szKey = InitKey(szKey) ' reset the encryption key
			i = 0			' reset the counter
		Else
			iCypherChar = 0
			For iPwr = 0 to 7 ' only ASCII test supported
				bKeyBit = GetBit(charKey, iPwr)
				bClearBit = GetBit(charClear, iPwr)
				If (bKeyBit XOr bClearBit) Then _
					iCypherChar = iCypherChar + (2 ^ iPwr)
			Next
			'WScript.echo charClear & " = " & CStr(iCypherChar)
			szEncrypt = szEncrypt & Chr(iCypherChar)
		End If

	Wend
End Function

Function GetBit(charToCheck, iBitToGet)
	Dim iTemp
	iTemp = Asc(charToCheck)
	iTemp = iTemp mod (2 ^ (iBitToGet + 1))
	If iTemp >= (2 ^ iBitToGet) Then
		GetBit = 1
	Else
		GetBit = 0
	End IF
End Function

Function InitKey(szKey)
	Dim i
	Randomize
	For i = 1 to 256
		szKey = szKey & Chr(Round((Rnd() * 94)) + 32)
	Next
	WshShell.RegWrite c_CryptKeyPath, szKey 
	dbgout "Writing key: " & szKey
	InitKey = szKey
End Function

' *************************
'  *** Network Functions ***
'   *************************

Function ActiveNetwork
	Dim bResult, objWMIService, objItem, colItems, objAddr
	bResult = False
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter Where NetConnectionStatus>0",,48)
	For Each objItem in colItems
		Select Case objItem.NetConnectionStatus
			Case 2
				DbgOut objItem.Description & "(NetConnectionStatus): " & "Connected"
				bResult = True
			Case 7: DbgOut objItem.Description & "(NetConnectionStatus): " & "Disconnected"
			Case Else: DbgOut objItem.Description & "(NetConnectionStatus): " & "Unknown (" & Str(objItem.NetConnectionStatus) & ")"
		End Select
	Next
	ActiveNetwork = bResult
End Function

' For information on MapNetworkDrive method, see: http://msdn.microsoft.com/library/en-us/script56/html/wsMthMapNetworkDrive.asp?frame=true
' For information on WshRemote Object, see http://msdn.microsoft.com/library/en-us/script56/html/wslrfRemote_WSHObject.asp?frame=true

Function Ping(szHost)
	' For info on Win32_PingStatus, see 
	' code adapted from Microsoft sample: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_tasks__networking.asp

	Dim objWMIService: Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colPings: Set colPings = objWMIService.ExecQuery _
	    ("Select * From Win32_PingStatus where Address = '" & szHost & "'")
	
	Dim objStatus
	For Each objStatus in colPings
	    If IsNull(objStatus.StatusCode) _
	        Or objStatus.StatusCode<>0 Then 
	        DbgOut "Ping(" & szHost & "): Computer did not respond." 
	    Else
	        DbgOut "Ping(" & szHost & "): Computer responded."
	        Ping = True
	    End If
	Next
	Set objWMIService = Nothing
	Set colPings = Nothing
	Set objStatus = Nothing
End Function

Function NetSend(szHost, szMessage)
	DbgOut "Sending message to " & szHost & ": """ & szMessage & """"
	NetSend = WshShell.Run("net send /DOMAIN:" & szHost & " Script[" & WScript.ScriptName & "]: " & szMessage,,True)
End Function

' *********************
'  *** Misc Routines ***
'   *********************

Function ReplaceSubstr(szSource, szFind, szReplace)
	Dim idx, iLastPos, szTemp
	
	ReplaceSubstr = szSource ' in case no replacement is necessary
	
	idx = InStr(ReplaceSubstr, szFind) ' set up the while loop var

	While idx > 0	' As long as szFind is found in ReplaceSubstr...
		iLastPos = Len(ReplaceSubstr) - Len(szFind)
		Select Case idx
			Case 1
				'szFind is at the beginning of target string
				ReplaceSubstr = szReplace & Right(ReplaceSubstr, iLastPos)
			Case iLastPos
				'szFind is at the end of target string
				ReplaceSubstr = Left(ReplaceSubstr, iLastPos) & szReplace
			Case Else
	 			'szFind is in the middle of target string
				szTemp = Mid(ReplaceSubstr, 1, idx - 1) & szReplace & _
						Mid(ReplaceSubstr, idx + Len(szFind))
				ReplaceSubstr = szTemp 
		End Select
		idx = InStr(ReplaceSubstr, szFind)
	Wend
End Function

Sub DbgOut(szMsg) 	' Echo messages in when Debug Mode is active.
	If bDebug Then
		If UCase(szMsg) = "QUIT" then 
			CriticalError "Debug Message """ + szMsg + """ causes script termination"
		End If
		Debug.Write szMsg
		If Not IsObject(oTSLog) Then
			OpenLogFile WScript.ScriptName 
		End If
		WriteLog "Debug: " & szMsg
	End If	
End Sub

Function CheckBit(iBitMask, iBitNum)
	' Implementation starts from LSB 
	' (0, 1), (1, 1) and (4, 1) return False and (2, 1), (3, 1) and (6, 1) return True
	Dim iScratch: iScratch = (iBitMask mod (2 ^ (iBitNum + 1)))
	CheckBit = (iScratch >= (2 ^ (iBitNum)))
End Function

Function EmptyArray(arrTest)
	On Error Resume Next
	Dim isize: isize = UBound(arrTest)
	Select Case Err
		Case 9
			EmptyArray = True
		Case 13
			EmptyArray = True
		Case Else
			CheckErrWarn "Checking array size", "EmptyArray"
	End Select
	Err.Clear
End Function
