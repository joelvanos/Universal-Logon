REM Version Info
	'  Universal Logon Script 1.9.2
	'  Updated 20160922
	'
	'
REM Source Info
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Source Found At http://hacks.oreilly.com/pub/h/1151#code
	' Source Author:   Dan Thomson, myITforum.com columnist
	'           I can be contacted at dethomson@hotmail.com
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'On Error Resume Next
REM Define Variables
	Dim objFileSys
	Dim objIntExplorer
	Dim objWshNetwork
	Dim objWshShell
	Dim objShell
	Dim objSysInfo
	Dim objWMIService
	Dim objNTUser
	Dim objFolder
	Dim objSubFolder
	Dim objAppData
	Dim objProgramFiles
	Dim objProgramFilesx86
	Dim objWindows
	Dim ObjUserProfile
	Dim objAUSM
	Dim objSM
	Dim objwshSystemEnv
	Dim objShortcut
	Dim objReg
	Dim colFiles
	Dim ObjFile
	Dim objWindow
	
	Dim colWMIResults
	Dim ItemWMIResults
	
	Dim arrKey
	Dim arrOSVersion
	Dim arrTemp
		
	Dim DicInstalledPrintersLocalExclude
	Dim DicMappedDrives
	Dim dicGroupList
	Dim DicInstalledPrinters
	
	Dim strIEVersion
	Dim strChromeVersion
	Dim strFireFoxVersion
	Dim strWorkstation 
	Dim strRealWorkstation      'Local Computer Name
	Dim strUserGroups       'List of groups the user is a member of
	Dim strFileServer		'Location of file server
	Dim strFileServerUsers	'Location of Users Files
	Dim strPrintOU			'Location of the groups for printers
	Dim strUserHomeShare	'Hold user home directory share
	Dim strPrintServer		'Location of Print server
	Dim strKey
	Dim strUserID
	Dim strComputerOU
	Dim StrInstalledPrintersLocal
	Dim StrOldDefault
	Dim StrForcePrinter
	Dim StrPaperVisionWebAssistant
	Dim StrLogUNC
	Dim StrShLib
	Dim StrTemp	
	Dim strComputerOULong
	Dim strOSVersion
	Dim StrCompanyDefaultWallpaper
	Dim StrCompanyDefaultWallpaperStyle
	Dim strCompanyName
	Dim StrContact
	
	Dim IntTemp	
	Dim intProcessorWidth
	Dim DateYMD
	
	Dim ExemptWorkstation
	Dim binChangeDefault
	Dim binPrintSpooler
	
REM Define Constants
	Const APPLICATION_DATA = &H1a
	Const PROGRAM_FILES = &H26
	Const PROGRAM_FILESX86 = &H2A       'x86 C:\Program Files on RISC
	Const USERPROFILE = &H28
	Const WINDOWS = &H24
	Const DESKTOP = &H0
	Const ALL_USERS_START_MENU = &H16	
	Const START_MENU = &Hb
	Const HKEY_CURRENT_USER = &H80000001 'HKEY_CURRENT_USER
	Const HKEY_LOCAL_MACHINE = &H80000002 'HKEY_LOCAL_MACHINE
	Const ForAppending = 8
	Const ForWriting = 2
REM Single Instance
	' Uses WMI to see of the script is running and if so exits.
	'Call SingleInstance()	
	'WScript.Sleep 3000
REM Initialize common scripting objects
	Set objFileSys    = CreateObject( "Scripting.FileSystemObject" )
	Set objShell 	  = CreateObject( "Shell.Application" )
	Set objWshNetwork = CreateObject( "WScript.Network" )
	Set objWshShell   = CreateObject( "WScript.Shell" )
	Set objReg        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
	Set objwshSystemEnv  = objWshShell.Environment( "User" )
REM Bind to dictionary object.
	Set objSysInfo   = CreateObject("ADSystemInfo")
	Set DicInstalledPrinters = CreateObject("Scripting.Dictionary")
	Set DicInstalledPrintersLocalExclude = CreateObject("Scripting.Dictionary")
	Set DicMappedDrives = CreateObject("Scripting.Dictionary")
	Set dicGroupList	= CreateObject("Scripting.Dictionary")
REM Gathers Ad user Object
	Set objNTUser = GetObject("LDAP://" & objSysInfo.UserName)
	strUserID = objNTUser.sAMAccountName
REM Get processor AddressWidth to know x64 or i386
	Set colWMIResults = objWMIService.ExecQuery("SELECT AddressWidth FROM Win32_Processor Where DeviceID='CPU0'")
		For Each ItemWMIResults in colWMIResults
			intProcessorWidth = ItemWMIResults.AddressWidth
		Next
REM Setup the Shell Folders 
	Set objFolder = objShell.Namespace(APPLICATION_DATA)
	Set objAppData = objFolder.Self
	Set objFolder = objShell.Namespace(PROGRAM_FILES)
	Set objProgramFiles = objFolder.Self
	If intProcessorWidth = "64" Then
		Set objFolder = objShell.Namespace(PROGRAM_FILESX86)
		Set objProgramFilesx86 = objFolder.Self
	End If
	Set objFolder = objShell.Namespace(WINDOWS)
	Set objWindows = objFolder.Self	
	Set objFolder = objShell.Namespace(ALL_USERS_START_MENU)
	Set objAUSM = objFolder.Self	
	Set objFolder = objShell.Namespace(START_MENU)
	Set objSM = objFolder.Self	
	Set objFolder = objShell.Namespace(USERPROFILE)
	Set ObjUserProfile = objFolder.Self		
	
REM Set Global Variables

	strCompanyName = "Company Name"
	strFileServer = "fp03"
	strFileServerUsers = "fp01"
	strPrintServer = "fp01"
	strPrintOU = "LDAP://OU=Printers,OU=Groups,DC=domain,DC=local"
	ExemptWorkstation = 0
	strUserGroups = ""
	strUserHomeShare =  "\\" & strFileServerUsers & "\users$\" & strUserID
	dicGroupList.CompareMode = vbTextCompare	
	binPrintSpooler = False
	StrForcePrinter= "PDFCREATOR"
	StrPaperVisionWebAssistant = "C:\Program Files (x86)\Digitech Systems\PaperVision\PVWA\DSI.PVWA.Host.exe"
	binChangeDefault = True
	StrLogUNC="\\fp01\logs$\sessions_csv"
	REM the Windows Dir is added at the end of the Setup the Shell Folders section
	StrCompanyDefaultWallpaper="system32\oobe\info\backgrounds\background1920x1200.jpg"
	StrCompanyDefaultWallpaperStyle="2"
	strComputerOULong = CStr(objSysInfo.ComputerName)
	StrCompanyDefaultWallpaper = objWindows.path & "\" & StrCompanyDefaultWallpaper
	StrContact = "Support Desk @ "
	StrShLib = objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe"
	DateYMD = DatePart("yyyy",Date) _
        & Right("0" & DatePart("m",Date), 2) _
        & Right("0" & DatePart("d",Date), 2)
	
REM 
	'List of Devices to exclude on local computers Put in Upper CASE
	'DicInstalledPrintersLocalExclude.Add objItem.PortName, objItem.Name
	DicInstalledPrintersLocalExclude.Add "SHRFAX","FAX"
	DicInstalledPrintersLocalExclude.Add "WEBEX DOCUMENT LOADER PORT","WEBEX DOCUMENT LOADER"
	DicInstalledPrintersLocalExclude.Add "PDFCMON", "PDFCREATOR"
	DicInstalledPrintersLocalExclude.Add "PORTPROMPT", "MICROSOFT PRINT TO PDF"
	DicInstalledPrintersLocalExclude.Add "MICROSOFT PRINT TO PDF", "MICROSOFT PRINT TO PDF"
	DicInstalledPrintersLocalExclude.Add "XPSPORT", "MICROSOFT XPS DOCUMENT WRITER"
	DicInstalledPrintersLocalExclude.Add "DOCUMENTS\*.PDF", "ADOBE PDF"
	DicInstalledPrintersLocalExclude.Add "FXC","FACSYS PRINTER"
	DicInstalledPrintersLocalExclude.Add "FAX","FAX"
	DicInstalledPrintersLocalExclude.Add "FACSYS FAX PRINTER","FXC"
	DicInstalledPrintersLocalExclude.Add "MICROSOFT XPS DOCUMENT WRITER", "XPSPORT"
	DicInstalledPrintersLocalExclude.Add "WEBEX DOCUMENT LOADER","WEBEX DOCUMENT LOADER PORT"
	DicInstalledPrintersLocalExclude.Add "SEND TO ONENOTE 2010","NUL"
	DicInstalledPrintersLocalExclude.Add "SEND TO ONENOTE 2013","NUL"
	DicInstalledPrintersLocalExclude.Add "SEND TO ONENOTE 16","NUL"
	DicInstalledPrintersLocalExclude.Add "SEND TO ONENOTE 2016","NUL"
	DicInstalledPrintersLocalExclude.Add "BROTHER QL-500","BROTHER QL-500"
	DicInstalledPrintersLocalExclude.Add "CANON GENERIC FAX DRIVER (FAX)","CANON GENERIC FAX DRIVER (FAX)"
	DicInstalledPrintersLocalExclude.Add "PDF-XCHANGE5","PDF-XCHANGE PRINTER 2012"
	DicInstalledPrintersLocalExclude.Add "HP EPRINT","HP EPRINT"

REM Set setting for IE
	Call IESettings
REM Check for error getting user-name
	If strUserID = "" Then
	  objWshShell.Popup "Logon script failed - " & StrContact, , _
	    "Logon script", 48
	  Call Cleanup
	End If
REM Calls Clock in Time
	' Bring up Clock in web page
	'If InGroup( objNTUser,Time") Then
	'	Call StartIE("https://timeclock","")
	'End IF
REM Calls IE for Logon script Echo
	Call StartIE("Logging", strCompanyName & " Desktop Configuration - Please Wait . . . ")
	'Get around Download Web Browser - Internet Explorer nag
	'https://support.microsoft.com/en-us/kb/3123303
	'for each objWindow in objShell.Windows
	'if InStr(objWindow.FullName,"iexplore") then
	'	if InStr(objWindow.document.title,"Download Web Browser - Internet Explorer") then
	'		objWindow.Quit
	'	end if
	'end if
	'next
	
REM Gather some basic system info
	Call GetSystemInfo
REM Logs Login time
	Call RecordLogon
REM Reset Citrix Receiver
	'Call ResetCitrixReceiver(20151113)
REM Display welcome message
	Call UserPrompt ("<H1>" & strCompanyName & " Desktop Configuration</H1><hr style=""width:100%""></hr>Welcome " _
		& objNTUser.FullName &" - <B>Please don't close this window.</B>" )
	Call UserPrompt ("You are logging on to <B>" & strRealWorkstation & "</B>." )	
	Call UserPrompt ("Current Date is: " & Date() & " at " & Time())
	'Add horizontal line as a 'break'
	objIntExplorer.Document.WriteLn("<hr style=""width:100%""></hr>")
REM  PrinterChange
	'Call PrinterRemove("5","Software" & False) 'If All is put in it will remove all network printers all the time. The Second command is to change print servers
REM Slow down the script
	'WScript.Sleep 1500
	'Call UserPrompt ("Computer OU: " & strComputerOU)
	'Call UserPrompt ("objSysInfo.ComputerName : " & objSysInfo.ComputerName)
REM Start Citrix Receiver 
	REM If intProcessorWidth = "64" Then
		REM If objFileSys.FileExists(objProgramFilesx86.path & "\Citrix\ICA Client\concentr.exe") Then
			REM Call UserPrompt ("Starting Citrix Receiver" )
			REM objWshShell.Run chr(34) & objProgramFilesx86.path & "\Citrix\ICA Client\concentr.exe" & chr(34) & " /startup", 0, False
		REM End If
	REM End If
	REM If objFileSys.FileExists(objAUSM.path & "\Programs\Startup\Receiver.lnk") Then
		REM objWshShell.Run chr(34) & objAUSM.path & "\Programs\Startup\Receiver.lnk" & chr(34), 0, False
	REM End If	
REM Maps people to the main printers Based on OU
	REM InStr add a white-list of OU's that are not having the printers mapped.	
	If (InStr(strComputerOULong,"Servers") = 0) Then
		REM if not a server map printer via groups
		Call PrinterGroupMapping(strPrintOU)
	End If
REM Maps People with local printers.	
	REM If InGroup( objNTUser,"SetLocalPrinterDefault" ) and strWorkstation = strRealWorkstation  Then
		REM 'force people in the SetLocalPrinterDefault to use the local default printer instead of group added printer
		REM If Not StrInstalledPrintersLocal = "" Then
			REM If (InStr(strComputerOULong,"Servers") = 0) Then
				REM objWshNetwork.SetDefaultPrinter StrInstalledPrintersLocal  
				REM Call UserPrompt ("<b> Set local default printer: " & StrInstalledPrintersLocal & "</b>")
				REM 'binChangeDefault = False
			REM End If
		REM End If 
	REM End If	

REM  Map drives, add shared printers and set default homepage based on computer name
	'Select Case UCase( strWorkstation )
		'Case "WT00806474B0B8"
			'ExemptWorkstation = True
			'Call AddPrinter (strPrintServer, "BR102", False, ExemptWorkstation)
			'Call AddPrinter (strPrintServer, "BR106", False, False)
		'Case Else			
	'End Select
	'Select Case UCase( strRealWorkstation )
	'	Case
		'Case Else
	'End Select
REM  This Section preforms actions based on User Name
	REM Select Case strUserID
		REM Case "user"
			REM Call MapDrive ("U:",strFileServer,"Shared\SFTP","SFTP")
		REM Case Else
	REM End Select
REM  This Section preforms actions based on Computer OU
	'If strComputerOU = "Training Room" Then
	'	ExemptWorkstation = True
	'	Call AddPrinter (strPrintServer, "BR137", False, ExemptWorkstation)
	'End If

REM  This section performs actions based on group membership
	If InGroup( objNTUser,"Shared-Drive" )  Then
		Call MapDrive ("S:",strFileServer,"Shared","Shared")
	End If 	
	REM If InGroup( objNTUser,"msgroup" )  Then
		REM Call UserPrompt ("Starting Paper Vision in the background")
			REM Call AddPrinter (strPrintServer, "ms01", True, True)
			REM If objFileSys.FileExists(StrPaperVisionWebAssistant) Then
				REM objWshShell.Run chr(34) & StrPaperVisionWebAssistant & chr(34), 0, False
			REM End If
	REM End If 
	
	'Maps Network Drives
	If InGroup( objNTUser,"Domain Users" )  Then
		REM Setup Favorites
		If Not objFileSys.FolderExists(strUserHomeShare & "\Favorites") Then
			If objFileSys.FolderExists(strUserHomeShare) Then
				objFileSys.CreateFolder(strUserHomeShare & "\Favorites")
			End If
		End IF	
		If objFileSys.FolderExists("I:\") Then
			objShell.NameSpace("I:\").Self.Name = "Home Drive"
		Else
			Call MapDrive ("I:", strFileServerUsers, "users$\" & strUserID,"Home Drive")
			objShell.NameSpace("I:\").Self.Name = "Home Drive"
		End If
		
		Call MapDrive ("T:",strFileServer,"Teams","Teams")
		Call MapDrive ("S:",strFileServer,"Shared","Shared")

		'Maps  Main Copiers
		If (InStr(strComputerOULong,"Servers") = 0) and (InStr(strComputerOULong,"Managed") = 0) Then
			Call AddPrinter (strPrintServer, "cc01", True, False)
			Call AddPrinter (strPrintServer, "cc02", True, False)
			Call AddPrinter (strPrintServer, "cc03", True, False)
			Call AddPrinter (strPrintServer, "cc04", True, False)
		End If
	End If
	
	If InGroup( objNTUser,"IT" ) or InGroup (objNTUser,"admin-accounts") Then	
		Call MapDrive ("T:",strFileServer,"Teams","Teams")
		Call MapDrive ("S:",strFileServer,"Shared","Shared")
		Call MapDrive ("R:",strFileServer,"Archive","Archive")
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add horizontal line as a 'break'
objIntExplorer.Document.WriteLn("<hr style=""width:100%""></hr>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Task-bar Setup	
	If Not InStr(strComputerOULong,"Citrix") > 0 or InStr(strComputerOULong,"Desktop") > 0 Then
		If UBound(arrOSVersion) = 2 Then
			If CInt(arrOSVersion(0) & arrOSVersion(1)) >= 6.1 Then
				REM Fix Libraries
				If intProcessorWidth = "64" Then
					If objFileSys.FileExists(objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe") Then
						If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
							Call UserPrompt ("Setting Up Libraries")
							StrTemp = chr(34) & objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe" & chr(34) & " remove " _
								& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
								& chr(34) & "c:\Users\Public\Documents" & chr(34)
							'Call UserPrompt ("Libraries Command: " & StrTemp)
							objWshShell.Run StrTemp, 0, False
							Err.Clear
						Else
							'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
						End If
					Else
						'Call UserPrompt ("ShLib Missing: " & objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe")	
						If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
							Call UserPrompt ("Setting Up Libraries")
							StrTemp = chr(34) & "\\" & objSysInfo.DomainDNSName & "\NETLOGON\SysinternalsSuite\ShLib.exe" & chr(34) & " remove " _
								& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
								& chr(34) & "c:\Users\Public\Documents" & chr(34)
							'Call UserPrompt ("Libraries Command: " & StrTemp)
							objWshShell.Run StrTemp, 0, False
							Err.Clear
						Else
							'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
						End If
								
					End If
				End If
				Call UserPrompt ("Setting Up Pinned Task-bar Items")
				REM Function PinItem(strlPath, strPin, blnRemove)
				'Remove Microsoft Store
				If  objFileSys.FileExists(objAUSM.path & "\Windows Store.lnk") Then
					Call UserPrompt ("Un-Pin Microsoft Store " & IntTemp & "  to Task-bar: " _ 
						& PinItem(objAUSM.path & "\Windows Store.lnk", "Taskbar", True))
				End If
				If  objFileSys.FileExists("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows Store.lnk") Then
					Call UserPrompt ("Un-Pin Microsoft Store " & IntTemp & "  to Task-bar: " _ 
						& PinItem("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows Store.lnk", "Taskbar", True))
				End If
				'Outlook, Excel
					Select Case intProcessorWidth
						Case 64
							'Outlook, Excel
							If  objFileSys.FolderExists(objProgramFilesx86.path & "\Microsoft Office") Then
								Set objFolder = objFileSys.GetFolder(objProgramFilesx86.path & "\Microsoft Office")
								For Each objSubFolder in objFolder.SubFolders
									If InStr(objSubFolder.name,"Office") = 1 Then
										If Right(objSubFolder.name,2) > IntTemp Then
											If objFileSys.FileExists(objSubFolder.path & "\OUTLOOK.EXE") Then
												IntTemp = Right(objSubFolder.name,2)
											End If
										End If
									End If
								Next
									
								If Not IntTemp = "" Then
									If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk") Then
										Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk", "Taskbar", False))
									Else
										Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objProgramFilesx86.path & "\Microsoft Office\Office" & IntTemp & "\OUTLOOK.EXE", "Taskbar", False))
									End If
									
									If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk") Then
										Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk", "Taskbar", False))
									Else
										Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objProgramFilesx86.path & "\Microsoft Office\Office" & IntTemp & "\Excel.EXE", "Taskbar", False))
									End If
								End If
							End If
							
							'IE
							arrTemp = Split(strIEVersion,".")
							'Verify existing Shortcut
							If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk") Then
								Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
								'Test and Fix bad shortcut
								If Instr(objShortcut.TargetPath,objProgramFilesx86.path) = 0 and Instr(objShortcut.WorkingDirectory,"%HOMEDRIVE%%HOMEPATH%") = 0 Then 
									'Versions older than 11 need to be set for the 32-bit version. IE 11 and newer need to be set for 64-bit version
									If cint(arrTemp(0)) < 11 Then
										Set objShortcut = Nothing
										objFileSys.deletefile objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer*.lnk"
										Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
										objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
										objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
										objShortcut.save()
										Set objShortcut = Nothing
									Else
										'Fix bad Working Directory
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
										'objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
										objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
										objShortcut.save()
										Set objShortcut = Nothing
									End If	
								Else
									REM Good Shortcut
									Set objShortcut = Nothing
										Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
										
								End If
							Else
								REM Pin IE if it was not pinned before
								
								If cint(arrTemp(0)) < 11 Then
									Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
										& PinItem(objProgramFilesx86.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))
									Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
									objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
									objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
									objShortcut.save()
									Set objShortcut = Nothing
								Else
									Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
										& PinItem(objProgramFiles.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))
									Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
									objShortcut.TargetPath = objProgramFiles.path & "\Internet Explorer\iexplore.exe"
									objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
									objShortcut.save()
									Set objShortcut = Nothing
								End If								
							End If
							Set arrTemp = nothing
							Set objShortcut = Nothing
							
							'Chrome
							If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk") Then
								Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
									& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk", "Taskbar", False))
							Else
								If objFileSys.FileExists(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk") Then
									Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
										& PinItem(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk", "Taskbar", False))
								End If	
							End If	

							'Windows Explorer
								If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
									Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
									If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
										objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
										objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
										objShortcut.WorkingDirectory = ""
										objShortcut.save()
										Set objShortcut = Nothing
									End If
									Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
										& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk", "Taskbar", False))
								Else
										If  objFileSys.FileExists(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk") Then
										
											Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
												& PinItem(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk", "Taskbar", False))
										End If
									If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then	
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
										If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
											objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
											objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
											objShortcut.WorkingDirectory = ""
											objShortcut.save()
											Set objShortcut = Nothing
										End If
									End If 
								End If	
								
						Case 32
							'Outlook, Excel
							If  objFileSys.FolderExists(objProgramFiles.path & "\Microsoft Office") Then
								Set objFolder = objFileSys.GetFolder(objProgramFiles.path & "\Microsoft Office")
								For Each objSubFolder in objFolder.SubFolders
									If InStr(objSubFolder.name,"Office") = 1 Then
										If Right(objSubFolder.name,2) > IntTemp Then
											If objFileSys.FileExists(objSubFolder.path & "\OUTLOOK.EXE") Then
												IntTemp = Right(objSubFolder.name,2)
											End If
										End If
									End If
								Next
									
								If Not IntTemp = "" Then
									If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk") Then
										Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk", "Taskbar", False))
									Else
										Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objProgramFiles.path & "\Microsoft Office\Office" & IntTemp & "\OUTLOOK.EXE", "Taskbar", False))
									End If
									
									If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk") Then
										Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk", "Taskbar", False))
									Else
										Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
											& PinItem(objProgramFiles.path & "\Microsoft Office\Office" & IntTemp & "\Excel.EXE", "Taskbar", False))
									End If
								End If
							End If
							
							'IE
							If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk") Then
								Call UserPrompt ("Pin Internet Explorer to Task-bar: " _ 
									& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
							Else
								Call UserPrompt ("Pin Internet Explorer to Task-bar: " _ 
									& PinItem(objProgramFiles.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))			
							End If
							
							'Chrome
							If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk") Then
								Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
									& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk", "Taskbar", False))
							Else
								If objFileSys.FileExists(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk") Then
									Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
										& PinItem(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk", "Taskbar", False))
								End If	
							End If
							
							'Windows Explorer
							If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
								Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
									& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk", "Taskbar", False))
									Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
									If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
										objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
										objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
										objShortcut.WorkingDirectory = ""
										objShortcut.save()
										Set objShortcut = Nothing
									End If
							Else
								If  objFileSys.FileExists(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk") Then
									Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
										& PinItem(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk", "Taskbar", False))
									If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
									If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
										objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
										objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
										objShortcut.WorkingDirectory = ""
										objShortcut.save()
										Set objShortcut = Nothing
									End If
									End If 
								End If
							End If	
						Case Else
					End Select	
				End If
								
				If Not InGroup (objNTUser,"wwt-admin-accounts") Then	
					'Remove Windows PowerShell
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
						Call UserPrompt ("Un-pin Windows PowerShell to Task-bar: " _ 
							& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk", "Taskbar", True))
					End If
						
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
						If  objFileSys.FileExists(objWindows.path & "\system32\WindowsPowerShell\v1.0\powershell.exe") Then
							Call UserPrompt ("Un-pin Windows PowerShell to Task-bar: " _ 
								& PinItem(objWindows.path & "\system32\WindowsPowerShell\v1.0\powershell.exe", "Taskbar", True))
						End If
					End If		
					
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
						objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk")			
					End If	
					
					'Remove Server Manager
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
						Call UserPrompt ("Un-pin Server Manager to Task-bar: " _ 
							& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk", "Taskbar", True))
					End If		
					
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
						If  objFileSys.FileExists(objWindows.path & "\system32\system32\ServerManager.msc") Then
							Call UserPrompt ("Un-pin Server Manager to Task-bar: " _ 
								& PinItem(objWindows.path & "\system32\system32\ServerManager.msc", "Taskbar", True))
						End If
					End If	
					
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
						objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk")	
					End If
						
					'Remove Windows Media Player
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
						Call UserPrompt ("Un-pin Windows Media Player to Task-bar: " _ 
							& PinItem(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk", "Taskbar", True))
					End If
					If  objFileSys.FileExists(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
						If  objFileSys.FileExists(objAUSM.path & "\Programs\Windows Media Player.lnk") Then
							Call UserPrompt ("Un-pin Windows Media Player to Task-bar: " _ 
								& PinItem(objAUSM.path & "\Programs\Windows Media Player.lnk", "Taskbar", True))
						End If
					If  objFileSys.FileExists(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
						objFileSys.DeleteFile(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk")	
					End If	
						
				End If	
				'Remove Duplicate names
				If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*(*).lnk") Then
					objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*(*).lnk")
				End If
			End If
		End If
	End If 
REM Shortcut Clean Up
	If Not InStr(strComputerOULong,"Citrix") > 0 or InStr(strComputerOULong,"Desktop") > 0 Then
		REM Call UserPrompt ("Shortcut Cleanup . . . ")
		REM If  objFileSys.FileExists(objProgramFilesx86.path & "\Citrix\ICA Client\SelfServicePlugin\SelfService.exe") Then
			REM objWshShell.Run Chr(34) & objProgramFilesx86.path & "\Citrix\ICA Client\SelfServicePlugin\SelfService.exe" & Chr(34) & " -logoff –rmPrograms -logon -poll -exit",1,True
		REM End If 
		REM If objFileSys.FolderExists(objSM.path) Then 
			REM Call UserPrompt ("&nbsp;&nbsp;&nbsp;Shortcut Cleanup:" & objSM.Name)
			REM ShortcutCleanUp  objSM.path
		REM End If 
		REM If objFileSys.FolderExists(strUserHomeShare & "\Desktop") Then 
			REM Call UserPrompt ("&nbsp;&nbsp;&nbsp;Shortcut Cleanup:" & "Desktop")
			REM ShortcutCleanUp  strUserHomeShare & "\Desktop"
		REM End If 
	End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add horizontal line as a 'break'
objIntExplorer.Document.WriteLn("<hr style=""width:100%""></hr>")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Setup Other Settings
	Call UserPrompt ("Setting Up Other Settings")
		objWshShell.RegWrite "HKCU\Software\VMware, Inc.\VMware Tools\ShowTray", 0 , "REG_DWORD"
	REM Sets the User name up for Microsoft Office
		objWshShell.RegWrite "HKCU\Software\Microsoft\Office\Common\UserInfo\UserName", objNTUser.FullName , "REG_SZ"
		
	REM Default Wallpaper Fix
		ChangeDefaultWallpater StrCompanyDefaultWallpaper,StrCompanyDefaultWallpaperStyle 
		
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Add horizontal line as a 'break'
		objIntExplorer.Document.WriteLn("<hr style=""width:100%""></hr>")
		' End section
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
REM Show Clock on Terminal Server
	'objReg.GetbinaryValue HKEY_CURRENT_USER,"Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2","Settings",arrKey
	'If isNull(arrKey) then
	'	arrKey = Array(&H28,&H0,&H0,&H0,&HFF,&HFF,&HFF,&HFF,&H2,&H0,&H0,&H0,&H3,&H0,&H0,&H0,&H3C,&H0,&H0,&H0,&H1E,&H0,&H0,&H0,&HFE,&HFF,&HFF,&HFF,&HFE,&H3,&H0,&H0,&H92,&H6,&H0,&H0,&H1C,&H4,&H0,&H0)
	'	objReg.CreateKey HKEY_CURRENT_USER,"Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2"
	'Else
	'	If Not arrKey(8) = 2 Then
	'		arrKey(8) = &H2	'"2" is on "A" is Off
	'	End If
	'End If
	'objReg.SetBinaryValue HKEY_CURRENT_USER,"Software\Microsoft\Windows\CurrentVersion\Explorer\StuckRects2","Settings",arrKey
	'Set arrKey = Nothing
REM Fix Adobe X Issue
	REM objReg.GetbinaryValue HKEY_CURRENT_USER,"Software\Adobe\Acrobat Reader\10.0\Privileged","bProtectedMode",arrKey
	REM If isNull(arrKey) then
		REM objReg.CreateKey HKEY_CURRENT_USER,"Software\Adobe\Acrobat Reader\10.0\Privileged"
	REM End If
	REM objWshShell.RegWrite "HKCU\Software\Adobe\Acrobat Reader\10.0\Privileged\bProtectedMode",0,"REG_DWORD"

REM Inform user that logon process is done -- Finished network log-on processes
	Call UserPrompt ("Finished network log-on processes")
	objIntExplorer.Quit( )
	Call Cleanup

	
'*********************************************************************************************** 	
'************************* End of Main Script Subs and Functions Below ************************* 	
'***********************************************************************************************	

REM Task-bar Setup
	Sub TaskBarSetup ()
 		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:     		TaskBarSetup
		' Purpose:  	Remove unwanted library-ms from users session along with pined taskbar items
		' Input:		
		' Output:
		' Dependencies	
		' Usage:		Call TaskBarSetup  
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		If Not InStr(strComputerOULong,"Citrix") > 0 or InStr(strComputerOULong,"Desktop") > 0 Then
			If UBound(arrOSVersion) = 2 Then
				If CInt(arrOSVersion(0) & arrOSVersion(1)) >= 6.1 Then
					REM Fix Libraries
					If intProcessorWidth = "64" Then
						If objFileSys.FileExists(objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe") Then
							If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
								Call UserPrompt ("Setting Up Libraries")
								StrTemp = chr(34) & objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe" & chr(34) & " remove " _
									& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
									& chr(34) & "c:\Users\Public\Documents" & chr(34)
								'Call UserPrompt ("Libraries Command: " & StrTemp)
								objWshShell.Run StrTemp, 0, False
								Err.Clear
							Else
								'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
							End If
						Else
							'Call UserPrompt ("ShLib Missing: " & objProgramFilesx86.path & "\SysinternalsSuite\ShLib.exe")	
							If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
								Call UserPrompt ("Setting Up Libraries")
								StrTemp = chr(34) & "\\" & objSysInfo.DomainDNSName & "\NETLOGON\SysinternalsSuite\ShLib.exe" & chr(34) & " remove " _
									& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
									& chr(34) & "c:\Users\Public\Documents" & chr(34)
								'Call UserPrompt ("Libraries Command: " & StrTemp)
								objWshShell.Run StrTemp, 0, False
								Err.Clear
							Else
								'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
							End If
									
						End If
					End If
					Call UserPrompt ("Setting Up Pinned Task-bar Items")
					REM Function PinItem(strlPath, strPin, blnRemove)
					'Remove Microsoft Store
					If  objFileSys.FileExists(objAUSM.path & "\Windows Store.lnk") Then
						Call UserPrompt ("Un-Pin Microsoft Store " & IntTemp & "  to Task-bar: " _ 
							& PinItem(objAUSM.path & "\Windows Store.lnk", "Taskbar", True))
					End If
					If  objFileSys.FileExists("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows Store.lnk") Then
						Call UserPrompt ("Un-Pin Microsoft Store " & IntTemp & "  to Task-bar: " _ 
							& PinItem("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Windows Store.lnk", "Taskbar", True))
					End If
					'Outlook, Excel
						Select Case intProcessorWidth
							Case 64
								'Outlook, Excel
								If  objFileSys.FolderExists(objProgramFilesx86.path & "\Microsoft Office") Then
									Set objFolder = objFileSys.GetFolder(objProgramFilesx86.path & "\Microsoft Office")
									For Each objSubFolder in objFolder.SubFolders
										If InStr(objSubFolder.name,"Office") = 1 Then
											If Right(objSubFolder.name,2) > IntTemp Then
												If objFileSys.FileExists(objSubFolder.path & "\OUTLOOK.EXE") Then
													IntTemp = Right(objSubFolder.name,2)
												End If
											End If
										End If
									Next
										
									If Not IntTemp = "" Then
										If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk") Then
											Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk", "Taskbar", False))
										Else
											Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objProgramFilesx86.path & "\Microsoft Office\Office" & IntTemp & "\OUTLOOK.EXE", "Taskbar", False))
										End If
										
										If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk") Then
											Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk", "Taskbar", False))
										Else
											Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objProgramFilesx86.path & "\Microsoft Office\Office" & IntTemp & "\Excel.EXE", "Taskbar", False))
										End If
									End If
								End If
								
								'IE
								arrTemp = Split(strIEVersion,".")
								'Verify existing Shortcut
								If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk") Then
									Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
									'Test and Fix bad shortcut
									If Instr(objShortcut.TargetPath,objProgramFilesx86.path) = 0 and Instr(objShortcut.WorkingDirectory,"%HOMEDRIVE%%HOMEPATH%") = 0 Then 
										'Versions older than 11 need to be set for the 32-bit version. IE 11 and newer need to be set for 64-bit version
										If cint(arrTemp(0)) < 11 Then
											Set objShortcut = Nothing
											objFileSys.deletefile objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer*.lnk"
											Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
											Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
											objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
											objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
											objShortcut.save()
											Set objShortcut = Nothing
										Else
											'Fix bad Working Directory
											Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
											'objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
											objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
											objShortcut.save()
											Set objShortcut = Nothing
										End If	
									Else
										REM Good Shortcut
										Set objShortcut = Nothing
											Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
											
									End If
								Else
									REM Pin IE if it was not pinned before
									
									If cint(arrTemp(0)) < 11 Then
										Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
											& PinItem(objProgramFilesx86.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
										objShortcut.TargetPath = objProgramFilesx86.path & "\Internet Explorer\iexplore.exe"
										objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
										objShortcut.save()
										Set objShortcut = Nothing
									Else
										Call UserPrompt ("Pin Internet Explorer " & cint(arrTemp(0)) & " to Task-bar: " _ 
											& PinItem(objProgramFiles.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk")
										objShortcut.TargetPath = objProgramFiles.path & "\Internet Explorer\iexplore.exe"
										objShortcut.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
										objShortcut.save()
										Set objShortcut = Nothing
									End If								
								End If
								Set arrTemp = nothing
								Set objShortcut = Nothing
								
								'Chrome
								If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk") Then
									Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
										& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk", "Taskbar", False))
								Else
									If objFileSys.FileExists(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk") Then
										Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
											& PinItem(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk", "Taskbar", False))
									End If	
								End If	

								'Windows Explorer
									If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
										If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
											objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
											objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
											objShortcut.WorkingDirectory = ""
											objShortcut.save()
											Set objShortcut = Nothing
										End If
										Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
											& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk", "Taskbar", False))
									Else
											If  objFileSys.FileExists(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk") Then
											
												Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
													& PinItem(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk", "Taskbar", False))
											End If
										If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then	
											Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
											If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
												objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
												objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
												objShortcut.WorkingDirectory = ""
												objShortcut.save()
												Set objShortcut = Nothing
											End If
										End If 
									End If	
									
							Case 32
								'Outlook, Excel
								If  objFileSys.FolderExists(objProgramFiles.path & "\Microsoft Office") Then
									Set objFolder = objFileSys.GetFolder(objProgramFiles.path & "\Microsoft Office")
									For Each objSubFolder in objFolder.SubFolders
										If InStr(objSubFolder.name,"Office") = 1 Then
											If Right(objSubFolder.name,2) > IntTemp Then
												If objFileSys.FileExists(objSubFolder.path & "\OUTLOOK.EXE") Then
													IntTemp = Right(objSubFolder.name,2)
												End If
											End If
										End If
									Next
										
									If Not IntTemp = "" Then
										If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk") Then
											Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Outlook.lnk", "Taskbar", False))
										Else
											Call UserPrompt ("Pin Outlook " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objProgramFiles.path & "\Microsoft Office\Office" & IntTemp & "\OUTLOOK.EXE", "Taskbar", False))
										End If
										
										If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk") Then
											Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Microsoft Office Excel.lnk", "Taskbar", False))
										Else
											Call UserPrompt ("Pin Excel " & IntTemp & "  to Task-bar: " _ 
												& PinItem(objProgramFiles.path & "\Microsoft Office\Office" & IntTemp & "\Excel.EXE", "Taskbar", False))
										End If
									End If
								End If
								
								'IE
								If objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk") Then
									Call UserPrompt ("Pin Internet Explorer to Task-bar: " _ 
										& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Internet Explorer.lnk", "Taskbar", False))
								Else
									Call UserPrompt ("Pin Internet Explorer to Task-bar: " _ 
										& PinItem(objProgramFiles.path & "\Internet Explorer\iexplore.exe", "Taskbar", False))			
								End If
								
								'Chrome
								If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk") Then
									Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
										& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Google Chrome.lnk", "Taskbar", False))
								Else
									If objFileSys.FileExists(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk") Then
										Call UserPrompt ("Pin Google Chrome to Task-bar: " _ 
											& PinItem(objAUSM.path & "\Programs\Google Chrome\Google Chrome.lnk", "Taskbar", False))
									End If	
								End If
								
								'Windows Explorer
								If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
									Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
										& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk", "Taskbar", False))
										Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
										If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
											objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
											objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
											objShortcut.WorkingDirectory = ""
											objShortcut.save()
											Set objShortcut = Nothing
										End If
								Else
									If  objFileSys.FileExists(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk") Then
										Call UserPrompt ("Pin Windows Explorer to Task-bar: " _ 
											& PinItem(objAUSM.path & "\Programs\Accessories\Windows Explorer.lnk", "Taskbar", False))
										If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" ) Then
											Set objShortcut = objWshShell.CreateShortcut(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Explorer.lnk" )
										If Not objShortcut.Arguments = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" Then
											objShortcut.TargetPath = "%SystemRoot%\explorer.exe"
											objShortcut.Arguments  = "/e,::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
											objShortcut.WorkingDirectory = ""
											objShortcut.save()
											Set objShortcut = Nothing
										End If
										End If 
									End If
								End If	
							Case Else
						End Select	
					End If
									
					If Not InGroup (objNTUser,"wwt-admin-accounts") Then	
						'Remove Windows PowerShell
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
							Call UserPrompt ("Un-pin Windows PowerShell to Task-bar: " _ 
								& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk", "Taskbar", True))
						End If
							
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
							If  objFileSys.FileExists(objWindows.path & "\system32\WindowsPowerShell\v1.0\powershell.exe") Then
								Call UserPrompt ("Un-pin Windows PowerShell to Task-bar: " _ 
									& PinItem(objWindows.path & "\system32\WindowsPowerShell\v1.0\powershell.exe", "Taskbar", True))
							End If
						End If		
						
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk") Then
							objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows PowerShell.lnk")			
						End If	
						
						'Remove Server Manager
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
							Call UserPrompt ("Un-pin Server Manager to Task-bar: " _ 
								& PinItem(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk", "Taskbar", True))
						End If		
						
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
							If  objFileSys.FileExists(objWindows.path & "\system32\system32\ServerManager.msc") Then
								Call UserPrompt ("Un-pin Server Manager to Task-bar: " _ 
									& PinItem(objWindows.path & "\system32\system32\ServerManager.msc", "Taskbar", True))
							End If
						End If	
						
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk") Then
							objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Server Manager.lnk")	
						End If
							
						'Remove Windows Media Player
						If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
							Call UserPrompt ("Un-pin Windows Media Player to Task-bar: " _ 
								& PinItem(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk", "Taskbar", True))
						End If
						If  objFileSys.FileExists(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
							If  objFileSys.FileExists(objAUSM.path & "\Programs\Windows Media Player.lnk") Then
								Call UserPrompt ("Un-pin Windows Media Player to Task-bar: " _ 
									& PinItem(objAUSM.path & "\Programs\Windows Media Player.lnk", "Taskbar", True))
							End If
						If  objFileSys.FileExists(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk") Then
							objFileSys.DeleteFile(objAUSM.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\Windows Media Player.lnk")	
						End If	
							
					End If	
					'Remove Duplicate names
					If  objFileSys.FileExists(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*(*).lnk") Then
						objFileSys.DeleteFile(objAppData.path & "\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\*(*).lnk")
					End If
				End If
			End If
		End If 
	
	End Sub
	
REM Libraries Cleanup
	Sub LibrariesCleanup ()
 		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:     		LibrariesCleanup
		' Purpose:  	Remove unwanted library-ms from users session
		' Input:		
		' Output:
		' Dependencies	intProcessorWidth,arrOSVersion,objFileSys,objProgramFilesx86,objAppData,objWshShell,StrShLib
		' Usage:		Call LibrariesCleanup  
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
		
			'Verify OS Compatibility 
			If UBound(arrOSVersion) = 2 Then
				If CInt(arrOSVersion(0) & arrOSVersion(1)) >= 6.1 Then
					REM Fix Libraries
					If intProcessorWidth = "64" Then
						'Look for ShLib.exe on the local computer 
						'Can be found in Microsoft Windows SDK for Windows 7
						If objFileSys.FileExists(StrShLib) Then
							If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
								Call UserPrompt ("Setting Up Libraries")
								'Remove Public Documents
								StrTemp = chr(34) & StrShLib & chr(34) & " remove " _
									& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
									& chr(34) & "c:\Users\Public\Documents" & chr(34)
								objWshShell.Run StrTemp, 0, False
								Err.Clear
							Else
								'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
							End If
						Else
							'Try to use ShLib.exe off of the NETLOGON share
							If objFileSys.FileExists("\\" & objSysInfo.DomainDNSName & "\NETLOGON\SysinternalsSuite\ShLib.exe") Then
								If objFileSys.FileExists( objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms") Then
									Call UserPrompt ("Setting Up Libraries")
									StrTemp = chr(34) & "\\" & objSysInfo.DomainDNSName & "\NETLOGON\SysinternalsSuite\ShLib.exe" & chr(34) & " remove " _
										& chr(34) & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms" & chr(34) & " " _ 
										& chr(34) & "c:\Users\Public\Documents" & chr(34)
									'Call UserPrompt ("Libraries Command: " & StrTemp)
									objWshShell.Run StrTemp, 0, False
									Err.Clear
								Else
									'Call UserPrompt ("Documents.library-ms Missing: " & objAppData.path & "\Microsoft\Windows\Libraries\Documents.library-ms")
								End If
							End If		
						End If
					End If
				End If
			End If
	End Sub
REM IE Settings
	Private Sub IESettings
 		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:     		IESettings
		' Purpose:  	Configure IE Zones and Pop-ups
		' Input:		
		' Output:
		' Dependencies		
		' Usage:		Call IESettings  
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'On Error Resume Next
	Rem Stops IE from checking for updates
		objWshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Main\NoUpdateCheck", 1 , "REG_DWORD"
			
	REM Adds Sites To the Exceptions List for Pop-ups
		objWshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\New Windows\Allow\res.cisco.com",0,"REG_BINARY"
		
	REM 'Add Sites to Local Intranet Zone	
	  REM 'Local Network
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range1\http",1, "REG_DWORD"
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range1\https",1, "REG_DWORD"
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range1\file",1, "REG_DWORD"
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges\Range1\:Range","10.68.*.*", "REG_SZ"
		REM 'File Server
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\fp01\file",1, "REG_DWORD"
	  REM 'Internal Domains
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\domain.local\http",1, "REG_DWORD"
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\domain.local\https",1, "REG_DWORD"
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\domain.local\file",1, "REG_DWORD"	
	REM 'Add Trusted Sites
		REM objWshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\cisco.com\res\https",2, "REG_DWORD"
		
		
 End Sub
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Shortcut Clean Up
	Private Sub ShortcutCleanUp (strFolder)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:     		ShortcutCleanUp
		' Purpose:  	Deletes Duplicate Shortcuts in Folder Recursively
		' Input:		strFolder
		' Output:
		' Dependencies		
		' Usage:		Call ShortcutCleanUp  [Folder with path]
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Dim ObjFile
		Dim ObjSubFolders
		Dim ObjSubFolder
		Dim objShortcut
		Dim colFiles
		Dim objFolder
		Dim objFileSys
		Dim dicpathList
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFileSys.GetFolder(strFolder)
		Set dicpathList	= CreateObject("Scripting.Dictionary")
		Set colFiles = objFolder.Files

		If colFiles.Count > 0 Then
			For Each ObjFile in objFolder.Files
				If lcase(Right(ObjFile.Name,3)) = "lnk" Then
					Set objShortcut = objWshShell.CreateShortcut(ObjFile.path)
					If dicpathList.Exists(objShortcut.Arguments) and len(objShortcut.Arguments) > 0 Then
						Call UserPrompt ("<font color=" & chr(34) & "red" & Chr(34) & ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & "Deleting Duplicate Shortcut: " & ObjFile.Name & "</font>")
						ObjFile.delete True
					Else
						If len(objShortcut.Arguments) > 0 Then dicpathList.Add objShortcut.Arguments,ObjFile.Name
					End If
				End If 
			Next
		End If 
		Set ObjSubFolders = objFolder.SubFolders
		If ObjSubFolders.count > 0 Then
			For Each ObjSubFolder in ObjSubFolders
				ShortcutCleanUp ObjSubFolder.path
			Next
		End If 
		 set colFiles = nothing
		 set ObjFile = nothing
		 set ObjSubFolders = nothing
		 set ObjSubFolder = nothing
		 set objFolder = nothing
		 set objFileSys = nothing
		 set dicpathList = nothing
	End Sub
REM Reset Citrix Receiver
	Private Sub ResetCitrixReceiver(IntDateChange)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:     		ResetCitrixReceiver
		' Purpose:  	Reset Citrix Receiver 4.3 +
		' Input:		IntDateChange
		' Output:
		' Dependencies		
		' Usage:		Call ResetCitrixReceiver  [date change]
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'On Error Resume Next
		Dim colProcessList
		Dim objProcess
		Dim objShell
		Dim intRCRDate
		Dim objWMIService
		Dim objRegistry
		Dim strKeyPath
		Dim strValueName
		Dim strValue
		Dim DResults
		Dim wmiLocator
		Dim wshNetwork
		Dim wmiNameSpace
		REM Gets Registry Value
		' Object used to get StdRegProv Namespace
		Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")

		' Object used to determine local machine name
		Set wshNetwork = CreateObject("WScript.Network")

		' Registry Provider (StdRegProv) lives in root\default namespace.
		Set wmiNameSpace = wmiLocator.ConnectServer(wshNetwork.ComputerName, "root\default")
		Set objRegistry = wmiNameSpace.Get("StdRegProv")
		'Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
		 
		strKeyPath = "Software\Citrix"
		strValueName = "RCRDate"
		objRegistry.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,strValue

		If IsNull(strValue) Then
			intRCRDate = 0
		Else
			intRCRDate = strValue
		End If

		REM DateYMD
		If IntDateChange > intRCRDate Then
			Call UserPrompt ("Resetting Citrix Receiver Last Reset: " & intRCRDate)
			REM Kill Citrix
			Set objWMIService = GetObject("winmgmts:" _
				& "{impersonationLevel=impersonate}!\\.\root\cimv2")
			Set colProcessList = objWMIService.ExecQuery _
				("Select * from Win32_Process Where Name = 'Receiver.exe' or Name = 'SelfServicePlugin.exe' or Name = 'SelfServicePlugin.exe' or Name = 'redirector.exe' or Name = 'wfcrun32.exe' or Name = 'concentr.exe' or Name = 'AuthManSvr.exe' or Name = 'SelfService.exe'")
			For Each objProcess in colProcessList
				objProcess.Terminate(1)
			Next
			REM Delete Citrix user key
			DResults = DeleteRegEntry(HKEY_CURRENT_USER,"Software\Citrix")
			REM Clean Roaming Appdata
			If objFileSys.FolderExists(objAppData.Path & "\Citrix") Then objFileSys.DeleteFolder (objAppData.Path & "\Citrix" ),True
			
			If objFileSys.FolderExists(objAppData.Path & "\ICAClient") Then objFileSys.DeleteFolder (objAppData.Path & "\ICAClient" ),True
			REM Clean Local Appdata
			If objFileSys.FolderExists(ObjUserProfile.Path & "\AppData\Local\Citrix") Then objFileSys.DeleteFolder (ObjUserProfile.Path & "\AppData\Local\Citrix" ),True
			REM Reset Receiver
			Set objShell = CreateObject("Wscript.Shell") 
			objShell.Run Chr(34) & objProgramFilesx86.path & "\Citrix\ICA Client\SelfServicePlugin\CleanUp.exe" & Chr(34) & " -cleanUser -silent"
			REM Update Registry
			objWshShell.RegWrite "HKCU\Software\Citrix\RCRDate", IntDateChange , "REG_DWORD"
			
			REM Clean up
			Set objShell = Nothing
			Set objProcess = Nothing
			Set colProcessList = Nothing
			Set objWMIService = Nothing
		End if 

	End Sub
REM DeleteRegEntry	
	Function DeleteRegEntry(sHive, sEnumPath)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Function:     DeleteRegEntry
		' Purpose:  	Recursively Deletes Registry Keys
		' Input:		
		'				sHive
		'				sEnumPath
		' Output:
		' Dependencies		
		' Usage:		Call DeleteRegEntry ( [file with full path],[ Style number}
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Dim objRegistry
		Dim wmiLocator
		Dim wshNetwork
		Dim wmiNameSpace
		Dim lRC
		Dim sNames
		Dim sKeyName
		
		' Object used to get StdRegProv Namespace
		Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")

		' Object used to determine local machine name
		Set wshNetwork = CreateObject("WScript.Network")

		' Registry Provider (StdRegProv) lives in root\default namespace.
		Set wmiNameSpace = wmiLocator.ConnectServer(wshNetwork.ComputerName, "root\default")
		Set objRegistry = wmiNameSpace.Get("StdRegProv")
		' Attempt to delete key.  If it fails, start the subkey
		' enumeration process.
		lRC = objRegistry.DeleteKey(sHive, sEnumPath)

		' The deletion failed, start deleting subkeys.
		If (lRC <> 0) Then

		' Subkey Enumerator
		   'On Error Resume Next

		   lRC = objRegistry.EnumKey(sHive, sEnumPath, sNames)

		   For Each sKeyName In sNames
			  If Err.Number <> 0 Then Exit For
			  lRC = DeleteRegEntry(sHive, sEnumPath & "\" & sKeyName)
		   Next

		   On Error Goto 0

		' At this point we should have looped through all subkeys, trying
		' to delete the registry key again.
		   lRC = objRegistry.DeleteKey(sHive, sEnumPath)

		End If
		Set wmiLocator = nothing
		Set objRegistry = nothing
		Set wshNetwork = nothing
		Set wmiNameSpace = nothing
		Set lRC = nothing
		Set sNames = nothing
		Set sKeyName = nothing
		
	End Function
			
				
REM Default Wallpaper Fix Sub
	Private Sub ChangeDefaultWallpater(StrlocCompanyDefaultWallpaper,StrlocCompanyDefaultWallpaperStyle)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	ChangeDefaultWallpater
		' Purpose:  	Overrides default wallpaper with your company's default wallpaper. Will not override user set wallpaper.
		' Input:		
		'				StrlocCompanyDefaultWallpaper file with full path
		'				StrlocCompanyDefaultWallpaperStyle Number
		' Output:
		' Dependencies	objFileSys	
		' Usage:		Call ChangeDefaultWallpater ( [file with full path],[ Style number}
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next
		Dim StrTemp
		Dim StrWallPaper
		
		REM Only change wallpaper if file exists.
		If objFileSys.FileExists(StrlocCompanyDefaultWallpaper) Then
			'Get current wallpaper
			StrWallPaper = objWshShell.RegRead("HKCU\Control Panel\Desktop\Wallpaper")
			Select Case Ucase(StrWallPaper)
				Case "C:\PROGRAM FILES\CITRIX\ENHANCEDDESKTOPEXPERIENCE\CITRIX_LOGO.JPG" 'XenApp Default Wallpaper
					objWshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", StrlocCompanyDefaultWallpaper , "REG_SZ"
					objWshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", StrlocCompanyDefaultWallpaperStyle , "REG_SZ"	
					Call RestartExplorer
				Case "C:\WINDOWS\WEB\WALLPAPER\WINDOWS\IMG0.JPG" 'Windows Default Wallpaper
					objWshShell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", StrlocCompanyDefaultWallpaper , "REG_SZ"
					objWshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", StrlocCompanyDefaultWallpaperStyle , "REG_SZ"	
					Call RestartExplorer
				Case Ucase(StrlocCompanyDefaultWallpaper)
					StrTemp = objWshShell.RegRead("HKCU\Control Panel\Desktop\WallpaperStyle")
					If Not StrTemp = StrlocCompanyDefaultWallpaperStyle Then 
						objWshShell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", StrlocCompanyDefaultWallpaperStyle , "REG_SZ"	
						Call RestartExplorer
					End If
				Case Else
			End Select
		End If
	End Sub
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Restart Explorer Sub 
	Private Sub RestartExplorer()
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	RestartExplorer
		' Purpose:  	Restarts Windows Explorer for settings to take effect.
		' Input:		
		' Output:
		' Dependencies		
		' Usage:		Call RestartExplorer 
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Dim objWMIService
		Dim colProcessList
		Dim objProcess
		Dim objShell
		' Kill Explorer.exe
		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colProcessList = objWMIService.ExecQuery _
			("Select * from Win32_Process Where Name = 'explorer.exe'")
		For Each objProcess in colProcessList
			objProcess.Terminate(1)
		Next
		Set objProcess = Nothing
		Set colProcessList = Nothing
		Set objWMIService = Nothing

		' Launch Explorer.exe
		Set objShell = CreateObject("Wscript.Shell") 
		objShell.Run "explorer.exe" 
		Set objShell = Nothing
	End Sub
  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Printer Change Sub
	Private Sub PrinterRemove(strRevision,StrUserRegPath,ChangePrinterServer)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	PrinterRemove
		' Purpose:  	Removes all dead and non strPrintServer Network Printers
		' Input:		strRevision Version number for registry
		'				StrUserRegPath Where to find the Revision
		'				ChangePrinterServer True/False
		' Output:
		' Dependencies		
		' Usage:
		'           	Call PrinterRemove ([Reg Revision Number], [Reg path], [New Printer Server])
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'On Error Resume Next
		
		Dim objPrinter
		Dim objNTPrinter
		Dim colInstalledPrinters
		Dim strDefaultPrinter
		Dim strKey
		Dim arrPrinter
		Dim objWMIService
		Dim objPrintServer
		Dim objPrintQueue
		Dim dicPrintServer
		Dim arrPrinterSplit
	
		Set dicPrintServer	= CreateObject("Scripting.Dictionary")
	
		'Gets list of printers on print server and add them to dicPrintServer
		Set objPrintServer = GetObject("WinNT://" & strPrintServer & ",Computer")
		objPrintServer.Filter = Array("PrintQueue")
		For Each objPrintQueue In objPrintServer
			arrPrinterSplit = Split(objPrintQueue.PrinterName,"\")
			dicPrintServer.Add arrPrinterSplit(3),arrPrinterSplit(2) 
		Next
		'Get default printer
		If ChangePrinterServer Then
			strDefaultPrinter = objWshShell.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows\Device")
			arrPrinter = Split(strDefaultPrinter,",")
			strDefaultPrinter = arrPrinter(0)
			arrPrinter = Split(strDefaultPrinter,"\")
			If UBound(arrPrinter) > 2 Then strDefaultPrinter = arrPrinter(3)
		End If
		'See if we have ran this before
		objReg.GetStringValue HKEY_CURRENT_USER,StrUserRegPath,"PrinterChange",strKey
		If isNull(strKey) Or strKey = "" or strKey <> strRevision or UCase(strRevision) = "ALL" Then
			Call UserPrompt ("Removing all Network Printers")
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
			Set colInstalledPrinters =  objWMIService.ExecQuery ("Select * from Win32_Printer")
			For Each objPrinter in colInstalledPrinters
				If Left(objPrinter.Name,2) = "\\" Then
					arrPrinter = Split(objPrinter.Name, "\")
					'If print server match 
					If UCase(arrPrinter(2)) = UCase( strPrintServer ) Then
						Set objNTPrinter = GetObject ("WinNT://" & arrPrinter(2) & "/" & arrPrinter(3))
						If IsObject( objNTPrinter ) AND (objNTPrinter.Name <> "" AND objNTPrinter.Class = "PrintQueue" )  And dicPrintServer.Exists(arrPrinter(3)) Then
							Call UserPrompt ("<b>Keeping Network Printer: " &  objPrinter.Name & "</b>")
						Else
							Call UserPrompt ("Removing Network Printer: " &  objPrinter.Name )
							objWshNetwork.RemovePrinterConnection objPrinter.Name
							DicInstalledPrinters.Remove(arrPrinter(3))
						End IF
					Else
						'If Printer server don't match
						If ChangePrinterServer Then
							If strDefaultPrinter = arrPrinter(3) Then
								Call AddPrinter (strPrintServer, arrPrinter(3),True,True)
							Else
								Call AddPrinter (strPrintServer, arrPrinter(3),True,False)
							End If
						End If
						Call UserPrompt ("Removing Network Printer: " &  objPrinter.Name )
						objWshNetwork.RemovePrinterConnection objPrinter.Name
						DicInstalledPrinters.Remove(arrPrinter(3))
					End IF
				End If
			Next
			
			If isNull(strKey) Or strKey = "" Then
				objReg.CreateKey HKEY_CURRENT_USER,StrUserRegPath & "\PrinterChange"
			End If 
			objReg.SetStringValue HKEY_CURRENT_USER,StrUserRegPath,"PrinterChange",strRevision
			'Add horizontal line as a 'break'
			objIntExplorer.Document.WriteLn("<hr style=""width:100%""></hr>")
		End If
	End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Get Main Printer Groups and Call AddPrinter Sub
	Sub PrinterGroupMapping(objOU)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	PrinterGroupMapping
		' Purpose:  	Maps Printers for users for all groups in OU
		' Input:		objOU - AD OU object
		' Output:
		' Dependencies	
			'Printer queue needs be be named the same as group
			'Printer server needs to be in the group description		
		' Usage:
		'           	Call PrinterGroupMapping ([AD OU Object])
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Dim objGroup
		Dim bnSetDefault
		Dim arrPrintDescription
		Dim strPrintqueue
		Dim strPrintServerLocal
		'On Error Resume Next

		Set objOU = GetObject(objOU)
		If Not IsEmpty (objOU) Then
			For Each objGroup in objOU	
				bnSetDefault = False
				Select Case objGroup.class
					Case "organizationalUnit"
						Call PrinterGroupMapping(objGroup.ADSPath)
					Case "group"
						If InGroup ( objNTUser,objGroup.sAMAccountName ) Then
							If Not objGroup.Get("description") = "" Then
								'Allows you to Added the word Default after the printer server to set that print as the Default
								'Setup Variables
								arrPrintDescription = split(objGroup.Get("description"), " ") 
								Select Case Ubound(arrPrintDescription)
									Case 0
										'Just has Printer Server
										strPrintServerLocal = arrPrintDescription(0)
										strPrintqueue = objGroup.sAMAccountName
									Case 1
										'Printer Server and Default
										strPrintServerLocal = arrPrintDescription(0)
										If Ucase(arrPrintDescription(1)) = "DEFAULT" Then
											bnSetDefault = True						
										End If
										strPrintqueue = objGroup.sAMAccountName
										
									Case 2
										'Printer Server , Default and Print Queue Name
										strPrintServerLocal = arrPrintDescription(0)
										If Ucase(arrPrintDescription(1)) = "DEFAULT" Then
											bnSetDefault = True						
										End If
										strPrintqueue = arrPrintDescription(2)
									Case Else
										strMsg = "Unable to connect to network printer. " & vbCrLf _
											& "Please contact: " & StrContact & vbCrLf _
											& "Ask them to check the " & objGroup.sAMAccountName & " printer group." _
											& vbCrLf _
											& "Let them know that you are unable to connect to the '" _
											& objGroup.sAMAccountName & "' printer description = " & objGroup.Get("description") _
											& vbCrLf & vbCrLf _
											& "Error: " & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
										objWshShell.Popup strMsg,, "Logon Error! Unable to connect to network printer.", 48
								End Select
								'Adds printer with printer server in the AD Group Description field.
								If binChangeDefault Then								
									Call AddPrinter (strPrintServerLocal, strPrintqueue, True, bnSetDefault)
								Else
									Call AddPrinter (strPrintServerLocal, strPrintqueue, True, False)
								End If 
							Else
								'Use Group name to map printer using default printer server
								If binChangeDefault Then
									'Try to use the default printer server to map printer
									Call AddPrinter (strPrintServer, objGroup.sAMAccountName, True, bnSetDefault)
								Else
									Call AddPrinter (strPrintServer, objGroup.sAMAccountName, True, False)
								End If	
							End If 
						End If
					Case Else
				End Select
			Next
		End If
	End Sub	
	
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Add Printer Sub
	Private Sub AddPrinter(strPrtServer, strPrtShare, blnNoError, blnSetDefaultPrinter)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	AddPrinter
		' Purpose:  	Connect to shared network printer
		' Input:
		'           	strPrtServer        Name of print server
		'           	strPrtShare         Share name of printer
		'				blnNoError				Disables Print Errors True/False
		'				blnSetDefaultPrinter	Set as Default Printer True/False
		' Output:
		' Dependencies	binPrintSpooler,DicInstalledPrinters
		' Usage:
		'           	Call AddPrinter ( "MyPrintServer", "SharedPrinter", False, True)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next

		Dim strPrtPath    'Full path to printer share
		Dim strMsg        'Message output to user
		Dim objNTPrinter
		Dim objWMIService
		Dim colItems, objItem
		Dim arrPrinterSplit
		Dim binMappedPrinter
		'Build path to printer share
		strPrtPath = "\\" & strPrtServer & "\" & strPrtShare
		binMappedPrinter = False

		'Printer Spooler is running then map printers
		If binPrintSpooler Then
			'Tests to see if the printer is already mapped.
			If Not DicInstalledPrinters.Exists(strPrtShare) Then
				'Getting Network Printer Object
				Set objNTPrinter = GetObject ("WinNT://" & strPrtServer & "/" & strPrtShare)
				'Test Printer Object
				If IsObject( objNTPrinter ) AND ( Not objNTPrinter.Name = "" AND objNTPrinter.Class = "PrintQueue")  Then
						strMsg = "Unable to connect to network printer. " & vbCrLf _
							& "Please contact: " & StrContact & vbCrLf _
							& "Ask them to check the " & strPrtServer & " server." _
							& vbCrLf & vbCrLf _
							& "Let them know that you are unable to connect to the '" _
							& strPrtShare & "' printer." _
							& vbCrLf & vbCrLf _
							& "Error: " & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
					'Show error if it is still having errors
					If  Not Err.Number = 0 Then
						If (blnNoError = False) and (Not Err.Source = "") Then
							objWshShell.Popup strMsg,, "Logon Error! Unable to connect to network printer.", 48
							err.clear
							Exit Sub
						Else
							Call UserPrompt (strMsg)
							err.clear
							Exit Sub
						End If
					End If
					objWshNetwork.AddWindowsPrinterConnection strPrtPath
					'Check error condition and output appropriate user message
					If Not Err.Number = 0 Then
						'WScript.Sleep 50
						Err.clear
						Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\CIMV2")
						'Verify that printer is mapped			
						Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Printer")
						If Not IsNull(colItems) Then
							For Each objItem In colItems
								If Left(objItem.Name,2) = "\\" Then
									If objItem.Name =  strPrtPath Then
										binMappedPrinter = True
									End IF
								End IF
							Next
							If binMappedPrinter Then
								'Adds printer to Dictionary
								DicInstalledPrinters.Add strPrtShare, strPrtShare	
								If blnSetDefaultPrinter Then
									'Set Printer as default
									Call UserPrompt ("<h3>Setting newly added default printer: <font color=" & chr(34) & "green" & chr(34) & ">" & strPrtShare & "</font></h3>")
									objWshNetwork.SetDefaultPrinter strPrtPath
								Else
									Call UserPrompt ("Successfully added printer connection to " & strPrtPath)	
								End If
							Else
								Call UserPrompt ( "Trying to map printer " & strPrtShare & " again")
								'Try Mapping Printer Again
								objWshNetwork.AddWindowsPrinterConnection strPrtPath
								strMsg = "Unable to connect to network printer. " & vbCrLf _
									& "Please contact: " & StrContact & vbCrLf _
									& "Ask them to check the " & strPrtServer & " server." _
									& vbCrLf & vbCrLf _
									& "Let them know that you are unable to connect to the '" _
									& strPrtShare & "' printer." _
									& vbCrLf & vbCrLf _
									& "Error: " & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
								'Show error if it is still having errors
								Call UserPrompt (strMsg)
								If  Not Err.Number = 0 Then
									If (Not Err.Source = "") and (blnNoError = False) Then
										objWshShell.Popup strMsg,, "Logon Error! Unable to connect to network printer.", 48
										err.clear
										Exit Sub
									Else
										Call UserPrompt (strMsg)
										err.clear
										Exit Sub									
									End If
								Else
									'Adds printer to Dictionary
									DicInstalledPrinters.Add strPrtShare, strPrtShare	
									If blnSetDefaultPrinter Then
										'Set Printer as default
										Call UserPrompt ("<h3>Setting newly added default printer on 2nd try: <font color=" & chr(34) & "green" & chr(34) & ">" & strPrtShare & "</font></h3>")
										objWshNetwork.SetDefaultPrinter strPrtPath
									Else
										Call UserPrompt ("Successfully added printer connection on 2nd try: " & strPrtPath)	
									End If
								End If
							End If
						End If						
						
					Else
						'Adds printer to Dictionary
						DicInstalledPrinters.Add strPrtShare, strPrtShare	
						If blnSetDefaultPrinter Then
							'Set Printer as default
							Call UserPrompt ("<h3>Setting newly added default printer: <font color=" & chr(34) & "green" & chr(34) & ">" & strPrtShare & "</font></h3>")
							objWshNetwork.SetDefaultPrinter strPrtPath
						Else
							Call UserPrompt ("Successfully added printer connection to " & strPrtPath)	
						End If
										
					End If	
				End if
				
			Else
				'Already connected to to printer
				If blnSetDefaultPrinter Then
					'Set Printer as default
					Call UserPrompt ("<h3>Setting already connected printer <font color=" & chr(34) & "green" & chr(34) & ">: " & strPrtShare & "</font> as default printer.</h3>")
					objWshNetwork.SetDefaultPrinter strPrtPath
				End If				
			End If 
		Else
			Call UserPrompt ("Printer Spooler Not running to connect: " & strPrtPath)
		End If 
		
	End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

REM MapDrive Sub
	Private Sub MapDrive( strDrive, strServer, strShare,strDriveName )
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	MapDrive
		' Purpose:  	Map a drive to a shared folder
		' Input:
		'           	strDrive    Drive letter to which share is mapped
		'           	strServer   Name of server that hosts the share
		'           	strShare    Share name
		'
		' Output:
		' Dependencies	DicMappedDrives
		' Usage:
		'           	Call MapDrive ("X:", "StaffServer", "StaffShare","Staff Share")
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next

		Dim strPath       'Full path to printer share
		Dim blnError      'True / False error condition
		Dim strMsg		  'Error Message 
		Dim objFileSys, objWshShell, objShell, objWshNetwork
		
		Set objFileSys    = CreateObject( "Scripting.FileSystemObject" )
		Set objShell 	  = CreateObject( "Shell.Application" )
		Set objWshNetwork = CreateObject( "WScript.Network" )
		Set objWshShell   = CreateObject( "WScript.Shell" )
		
		blnError = False
		'Build path to share
		strPath = "\\" & strServer & "\" & strShare
		'Call UserPrompt (strpath)

		'Check to see if the mapped drive is the same as the one we are trying to map.
		If DicMappedDrives.Exists(strDrive) Then
			If Not DicMappedDrives.Item(strDrive) = strPath Then
				'Disconnect Drive if drive letter is already mapped.
				'This assures everyone has the same drive mappings
				objWshNetwork.RemoveNetworkDrive strDrive, , True
				If Not Err.Number = 0 Then 
					objWshShell.Popup "Network Mapping Error: Error: "   _
					& vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source,, "Logon Error !", 48
				End If 
				
				'Test to see if share exists. Proceed if yes, set error condition if no.
				If objFileSys.FolderExists(strPath) Then
					Err.Clear
					'To get around a issue with Windows Vista/7 drive mapping issues need to make persistent connection
					objWshNetwork.MapNetworkDrive strDrive, strPath, True
					If Err.Number = 0 And objFileSys.DriveExists(strDrive & "\") Then
						DicMappedDrives.Item(strDrive) = strPath
					End If 
				Else
					blnError = True
				End If
				' Test if drive mapped.
				If objFileSys.DriveExists(strDrive) = False Then
					blnError = True
				End If
				'Check error condition and output appropriate user message
				If Err.Number <> 0 And blnError = True Then
					'Display message box informing user that the connection failed
					strMsg = "Unable to connect to network share. " & vbCrLf & _
							 "Please contact:  " & StrContact  & vbCrLf & _
							 "Ask them to check the " & strServer & " server." & vbCrLf & _
							 "Let them know that you are unable to connect to the " & _
							 "'" & strPath & "' share" & _
							 vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
					objWshShell.Popup strMsg,, "Logon Error Re-Mapping Drive!", 48
				Else
					If Not strDriveName = "" Then
						objShell.NameSpace(strDrive).Self.Name = strDriveName
					End If
					Call UserPrompt ("Successfully re-mapped drive connection to " & strPath & " ( " & strDrive  & " "& strDriveName & " ) " )
				End If

			End If 
		Else
			'If the drive does not exists map it.
			'Test to see if share exists. Proceed if yes, set error condition if no.
			If Not objFileSys.DriveExists(strDrive & "\") Then
				If objFileSys.FolderExists(strPath) Then
					'To get around a issue with Windows Vista/7 drive mapping issue need to make persistent connection
					objWshNetwork.MapNetworkDrive strDrive, strPath, True
					Select Case Err.Number
					Case 0
						If Not strDriveName = "" Then
							objShell.NameSpace(strDrive).Self.Name = strDriveName
						End If
						Call UserPrompt ("Successfully added mapped drive connection to " & strPath & " ( " & strDrive  & " "& strDriveName & " ) " )
						If objFileSys.DriveExists(strDrive & "\") Then DicMappedDrives.Add  strDrive, strPath				
					Case -2147024865, -2147024832, -2147024811, -2147023694, 80070055 
						Call UserPrompt ("Drive already mapped to " & strPath & " ( " & strDrive  & " "& strDriveName & " ) " )
					Case Else
						blnError = True
						'Display message box informing user that the connection failed
						strMsg = "Unable to connect to network share. " & vbCrLf & _
								 "Please contact:  " & StrContact  & vbCrLf & _
								 "Ask them to check the " & strServer & " server." & vbCrLf & _
								 "Let them know that you are unable to connect to the " & _
								 "'" & strPath & "' share" & _
								 vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
						objWshShell.Popup strMsg,, "Logon Error Mapping Drive!", 48
					End Select 
					

				Else
					strMsg = "Unable to connect to network share. " & vbCrLf & _
							 "Please contact:  " & StrContact  & vbCrLf & _
							 "Ask them to check the " & strServer & " server." & vbCrLf & _
							 "Let them know that you are unable to connect to the " & _
							 "'" & strPath & "' share" & _
							 vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
					objWshShell.Popup strMsg,, "Logon Error Share Does Not Exists!", 48
				End If
				
			Else
				strMsg = "Unable to connect to network share. " & vbCrLf & _
						 "Please contact:  " & StrContact  & vbCrLf & _
						 "Ask them to check the " & strServer & " server." & vbCrLf & _
						 "Let them know that you are unable to connect to the " & _
						 "'" & strPath & "' share" & _
						 vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
				objWshShell.Popup strMsg,, "Logon Error Drive Already Mapped!", 48
				If objFileSys.FolderExists(strDrive) Then
						blnError = False
					Else 
						Err.Clear
						objWshNetwork.MapNetworkDrive strDrive, strPath
						Select Case Err.Number
						Case 0
							If Not strDriveName = "" Then
								objShell.NameSpace(strDrive).Self.Name = strDriveName
							End If
							Call UserPrompt ("Successfully added mapped drive connection to " & strPath & " ( " & strDrive  & " "& strDriveName & " ) " )

							If Err.Number = 0 Then
								DicMappedDrives.Add  strDrive, strPath
							End If 					
						Case -2147024865, -2147024832, -2147024811, 80070055 
							Call UserPrompt ("Drive already mapped to " & strPath & " ( " & strDrive  & " "& strDriveName & " ) " )
						Case Else
							blnError = True
							'Display message box informing user that the connection failed
							strMsg = "Unable to connect to network share. " & vbCrLf & _
									 "Please contact:  " & StrContact  & vbCrLf & _
									 "Ask them to check the " & strServer & " server." & vbCrLf & _
									 "Let them know that you are unable to connect to the " & _
									 "'" & strPath & "' share" & _
									 vbCrLf & Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
							objWshShell.Popup strMsg,, "Logon Error 2nd Try!", 48
						End Select
					End If
			End if 
		End If

	End Sub
REM PinItem	
	Function PinItem(strlPath, strPin, blnRemove)
	'********************************************************************
	'* Function PinItem()
	'* Purpose:  Pin item to the Start Menu.
	'* Input:          strlPath         Path of exe to pin
	'*                 strPin        	Pin item to strPin
	'*                 					Values:
	'*										Taskbar		 (At least Windows 7)
	'*										Start Menu	 (At least Windows 7)
	'*										Quick Access (At least Windows 10)
	'*
	'* Dependencies:   objShell         Shell.Application object
	'*                 objFileSys       File System object
	'*				   arrOSVersion		Array that has Windows Version
	'* Returns:        True if the shortcut is created, else false
	'********************************************************************
		'On Error Resume Next

		Dim colVerbs
		Dim itemverb
		Dim blnQuickAccess
		Dim objFolder
		Dim objFolderItem
		Dim strFolder
		Dim strFile

		blnQuickAccess = False
		
		If UBound(arrOSVersion) = 2 Then
			If CInt(arrOSVersion(0) & arrOSVersion(1)) >= 6.1 Then
				If CInt(arrOSVersion(0) & arrOSVersion(1)) >= 10 Then blnQuickAccess = True		
				'***** Do nothing,Correct Version
			Else
				Call UserPrompt ( "This version of Windows does not support Pined Items.")
				Exit Function
			End If
		Else
			Call UserPrompt ( "Could not get Windows version.")
			Exit Function
		End If
		
		If objFileSys.FileExists(strlPath) Then
			'***** Do nothing, folder exists
		Else
			'***** Folder does not exist
			PinItem = False
			Call UserPrompt ( "File to pin does not exist.")
			Call UserPrompt ( "Please check the input and try again.")
			
			Exit Function
		End If

		strFolder = objFileSys.GetParentFolderName(strlPath)
		strFile = objFileSys.GetFileName(strlPath)
		'Call UserPrompt ( "Folder: " & strFolder )
		'Call UserPrompt ( "File: " & strFile)
		Err.Clear
		Set objFolder = objShell.Namespace(strFolder)
		Set objFolderItem = objFolder.ParseName(strFile)

		' ***** InvokeVerb for this does not work on Vista/WS2008
		'objFolderItem.InvokeVerb("P&in to Start Menu")

		' ***** This code works on Vista/WS2008+
		Set colVerbs = objFolderItem.Verbs

		Select Case Ucase(strPin)
			Case "TASKBAR"
				For each itemverb in objFolderItem.verbs
					If blnRemove Then
						If Replace(itemverb.name, "&", "") = "Unpin from Taskbar" Then itemverb.DoIt
					Else
						If Replace(itemverb.name, "&", "") = "Pin to Taskbar" Then itemverb.DoIt
					End if	
			   Next 
			Case "START MENU"
				For each itemverb in objFolderItem.verbs
					If blnRemove Then
						If Replace(itemverb.name, "&", "") = "Unpin from Start Menu" Then itemverb.DoIt
					Else
						If Replace(itemverb.name, "&", "") = "Pin to Start Menu" Then itemverb.DoIt
					End If	
				Next 
			Case "QUICK ACCESS" 'Windows 10 --Needs more work to prefect
				For each itemverb in objFolderItem.verbs
					If blnRemove Then
						If Replace(itemverb.name, "&", "") = "Unpin from Quick Access" Then itemverb.DoIt
						If Replace(itemverb.name, "&", "") = "Remove from Quick Access" Then itemverb.DoIt
					Else
						If Replace(itemverb.name, "&", "") = "Pin to Quick Access" Then itemverb.DoIt
					End If	
				Next 			
			Case Else
			'***** Do nothing
		End Select

		If Err.Number = 0 Then
			PinItem = True
		Else
			PinItem = False
		End If
	End Function
	
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM InGroup	
	Function InGroup(objADObject, strGroup)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Function:     InGroup
		' Purpose:  	Gets AD Users Groups and caches them in dicGroupList
		' Input:		objPriADObject,objSubADObject
		' Output:   	Set: dicGroupList
		' Dependencies	dicGroupList
		' Usage:    	Call InGroup [AD Object] , [Group Name]
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If (dicGroupList.Exists(objADObject.sAMAccountName & "\") = False) Then
			Call LoadGroups(objADObject, objADObject)
			dicGroupList.Add objADObject.sAMAccountName & "\", True
		End If
		InGroup = dicGroupList.Exists(objADObject.sAMAccountName & "\" & strGroup)
	End Function
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM LoadGroups
	Sub LoadGroups(objPriADObject, objSubADObject)
		'Taken from http://www.rlmueller.net/Programs/Logon2.txt
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	LoadGroups
		' Purpose:  	Gets AD Users for a Group with groups as a memberOf
		' Input:		objPriADObject,objSubADObject
		' Output:   	Set: dicGroupList
		' Dependencies	dicGroupList
		' Usage:    	Call LoadGroups [AD Object] , [Other AD Object]
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		Dim colstrGroups, objGroup, strCounter
		
		colstrGroups = objSubADObject.memberOf
		
		Select Case True
			Case IsEmpty(colstrGroups) = True
				Exit Sub
			Case TypeName(colstrGroups) = "String"
				Set objGroup = GetObject("LDAP://" & colstrGroups)
				If (dicGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
					dicGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
					Call LoadGroups(objPriADObject, objGroup)
				End If
				Exit Sub
			Case Else
				For strCounter = 0 To UBound(colstrGroups)
					Set objGroup = GetObject("LDAP://" & colstrGroups(strCounter))
					If (dicGroupList.Exists(objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName) = False) Then
						dicGroupList.Add objPriADObject.sAMAccountName & "\" & objGroup.sAMAccountName, True
						Call LoadGroups(objPriADObject, objGroup)
					End If
				Next
		End Select
	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

REM GetSystemInfo
	Private Sub GetSystemInfo
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	GetSystemInfo
		' Purpose:  	Gather basic information about the local system
		' Input:
		' Output:   	Set: strOSVersion,strWorkstation,strRealWorkstation,strComputerOU,binPrintSpooler,DicInstalledPrinters,DicMappedDrives,strIEVersion,strChromeVersion,strFireFoxVersion,StrOldDefault
		' Dependencies	objWshShell,objWshNetwork,objWMIService,DicInstalledPrinters,DicInstalledPrintersLocalExclude,DicMappedDrives
		' Usage:    	Call GetSystemInfo
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next
		Dim strName
		Dim arrDNComputer
		Dim arrPrinterSplit
		Dim colItems
		Dim objItem
		Dim objDrives
		Dim i
		Dim colRunningServices
		Dim objService 
		Dim StrLastNetworkPrinter
		Dim getOSVersion
		
		'Gets OS version
		'Should find out how to do this via WMI
		Set getOSVersion = objWshShell.exec("%comspec% /c ver")	
		Do Until getOSVersion.Status
			Wscript.Sleep 250
		Loop
		
		strOSVersion  = Trim(getOSVersion.stdout.readall)
		'Src String    Start on n in Version and add 1 for the space   
		'Get the length and subtract the starting position mince 3 (n + space + ])
		strOSVersion  = Mid(strOSVersion,InStr(strOSVersion,"n ") + 1,Len(strOSVersion) - (InStr(strOSVersion,"n ") +3) )
		arrOSVersion = Split (strOSVersion,".")
		
		'Get computer or terminal server name
		strName = objWshShell.ExpandEnvironmentStrings( "%CLIENTNAME%" )
		If strName <> "%CLIENTNAME%" AND strName <> "" Then
			'Set strWorkstation to the real name and not the name of the server
			strWorkstation = objWshShell.ExpandEnvironmentStrings( "%CLIENTNAME%" )
		Else
			strWorkstation = objWshNetwork.ComputerName
		End If
		strRealWorkstation = objWshNetwork.ComputerName

		'Gets the OU the Computer is in AD
		arrDNComputer = Split(objSysInfo.ComputerName,",")
		strComputerOU = Right (arrDNComputer(1),len(arrDNComputer(1)) - 3)		
		
		'Check to see if the Printer Spooler is running
		Set colRunningServices =  objWMIService.ExecQuery("Select * from Win32_Service Where Name = 'Spooler'")
		For Each objService in colRunningServices 
			If objService.State = "Running" Then binPrintSpooler = True
		Next
		
		'Makes a Dictionary of all currently installed printers; if printer spooler is running		
		Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Printer")
		If Not IsNull(colItems) and binPrintSpooler Then
			For Each objItem In colItems
				'Gets current default
				If objItem.Default = True Then
					StrOldDefault = objItem.Name
				End if					
				If Left(objItem.Name,2) = "\\" Then 'Only add network printers to DicInstalledPrinters
					If Not DicInstalledPrinters.Exists(objItem.ShareName) Then
						'Added Printer to DicInstalledPrinters id not already in DicInstalledPrinters
						arrPrinterSplit = split(objItem.Name,"\")
						DicInstalledPrinters.Add objItem.ShareName, arrPrinterSplit(2)
						StrLastNetworkPrinter = objItem.Name
					End if 
				End If
				If objItem.SystemName = strRealWorkstation Then	'Check for a remote session 				
					If DicInstalledPrintersLocalExclude.Exists(Ucase(objItem.PortName)) or DicInstalledPrintersLocalExclude.Exists(Ucase(objItem.PortName & ":")) or DicInstalledPrintersLocalExclude.Exists(Ucase(objItem.Name)) Then
						'Excluded local printers
					else	
						'Make sure local printer is not a fax
						If Instr(objItem.Name,"(FAX)") = 0 Then
							StrInstalledPrintersLocal = objItem.Name	
						End If
					End If
				End If 
			Next
		End If
		
		'Makes a Dictionary of all currently mapped drives	
		Set objDrives = objWshNetwork.EnumNetworkDrives		
		For i = 0 to objDrives.Count - 1 Step 2
		    If Not DicMappedDrives.Exists(objDrives.Item(i)) Then 'Add only new items do DicMappedDrives
				DicMappedDrives.Add  objDrives.Item(i), objDrives.Item(i+1)
		    End If
		Next
		'Resets Default printer off DicInstalledPrintersLocalExclude List
		'Call UserPrompt ("StrOldDefault=" & StrOldDefault )
		REM If DicInstalledPrintersLocalExclude.Exists(Ucase(StrOldDefault)) Then
			REM If Not StrInstalledPrintersLocal = "" Then
				REM objWshNetwork.SetDefaultPrinter StrInstalledPrintersLocal
				REM Call UserPrompt ("Changing Printer to: <B>" & StrInstalledPrintersLocal & " </B> from: " & StrOldDefault)
			REM Elseif Not StrLastNetworkPrinter = "" Then
				REM objWshNetwork.SetDefaultPrinter StrLastNetworkPrinter
				REM Call UserPrompt ("Changing Printer to: <B>" & StrLastNetworkPrinter & " </B> from: " & StrOldDefault)				
			REM Else 
				REM 'objWshNetwork.SetDefaultPrinter StrForcePrinter
				REM 'Call UserPrompt ("Changing Printer to: <B>" & StrForcePrinter & " </B> from: " & StrOldDefault)					
			REM End If
		REM End If
		
		REM Web Browser Versions
			'Chrome Version
			If Not Isnull(objProgramFilesx86) Then
				If objFileSys.FileExists(objProgramFilesx86.path & "\Google\Chrome\Application\chrome.exe") Then
					strChromeVersion = objFileSys.GetFileVersion(objProgramFilesx86.path & "\Google\Chrome\Application\chrome.exe")		
				End If	
			End If
			If objFileSys.FileExists(objProgramFiles.path & "\Google\Chrome\Application\chrome.exe") Then
				strChromeVersion = objFileSys.GetFileVersion(objProgramFiles.path & "\Google\Chrome\Application\chrome.exe")
			End If
			'FireFox Version
			If Not Isnull(objProgramFilesx86) Then
				If objFileSys.FileExists(objProgramFilesx86.path & "\Mozilla Firefox\firefox.exe") Then
					strFireFoxVersion = objFileSys.GetFileVersion(objProgramFilesx86.path & "\Mozilla Firefox\firefox.exe")		
				End If
			End If
			If objFileSys.FileExists(objProgramFiles.path & "\Mozilla Firefox\firefox.exe") Then
				strFireFoxVersion = objFileSys.GetFileVersion(objProgramFiles.path & "\Mozilla Firefox\firefox.exe")
			End If

			'IE Version
			If objFileSys.FileExists(objProgramFiles.path & "\Internet Explorer\iexplore.exe") Then
				strIEVersion = objFileSys.GetFileVersion(objProgramFiles.path & "\Internet Explorer\iexplore.exe")
			End If
		 REM If strChromeVersion Then Call UserPrompt ("Chrome Version: " & strChromeVersion)
		 REM If strIEVersion Then Call UserPrompt ("IE Version: " & strIEVersion)
		 
		Set strName					= Nothing
		Set arrDNComputer			= Nothing
		Set arrPrinterSplit			= Nothing
		Set colItems				= Nothing
		Set objItem					= Nothing
		Set objDrives				= Nothing
		Set i						= Nothing
		Set colRunningServices		= Nothing
		Set objService 				= Nothing
		Set StrLastNetworkPrinter	= Nothing
		Set getOSVersion			= Nothing
	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM RecordLogon Sub
	Private Sub RecordLogon()
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	RecordLogon
		' Purpose:  	Records user logon Name,Computer,Date, Time,Action
		' Input:
		' Output:
		' Dependencies	StrLogUNC,objFileSys
		' Usage:    	Call RecordLogon
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next
		
		Dim objFile
		
		If Not objFileSys.FileExists(StrLogUNC & "\" & strUserID & ".csv") Then
			Set objFile = objFileSys.OpenTextFile(StrLogUNC & "\" & strUserID & ".csv", ForWriting, True)
			objFile.WriteLine("UserName,ComputerName,Date,Time,Status")
			objFile.WriteLine(strUserID & "," & strRealWorkstation & "," & Date & "," & Time & ",Logging On")
			objFile.Close
			If Not Err.Number = 0 Then
				Set objFile = objFileSys.OpenTextFile(StrLogUNC & "\" & strUserID & ".csv", ForWriting, True)
				objFile.WriteLine("UserName,ComputerName,Date,Time,Status")
				objFile.WriteLine(strUserID & "," & strRealWorkstation & "," & Date & "," & Time & ",Logging On")
				objFile.Close
				Err.clear
			End If 
		Else
			Set objFile = objFileSys.OpenTextFile(StrLogUNC & "\" & strUserID & ".csv", ForAppending, True)
			objFile.WriteLine(strUserID & "," & strRealWorkstation & "," & Date & "," & Time & ",Logging On")
			objFile.Close
			If Not Err.Number = 0 Then
				Set objFile = objFileSys.OpenTextFile(StrLogUNC & "\" & strUserID & ".csv", ForAppending, True)
				objFile.WriteLine(strUserID & "," & strRealWorkstation & "," & Date & "," & Time & ",Logging On")
				objFile.Close	
				Err.clear
			End If 				
		End if
	
		Err.clear
	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Setup IE Sub
	Private Sub StartIE(strPage,strTitle)
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      SetupIE
		' Purpose:  Set up Internet Explorer for use as a status message window
		' Input:	strPage,strTitle
		' Output:	objIntExplorer
		' Dependencies 
		' Usage:    Call SetupIE [URL or "LOGGING"],[Window Title]
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		On Error Resume Next
		Dim intCount    'Counter used during AppActivate

		'Create reference to objIntExplorer
		'This will be used for the user messages. Also set IE display attributes
		If Not isnull(objIntExplorer) Then Set objIntExplorer = Nothing
		Set objIntExplorer = Wscript.CreateObject("InternetExplorer.Application")
		If Err.Number <> 0 Then
			WScript.Sleep(5000)
			if Not isnull(objIntExplorer) Then Set objIntExplorer = Nothing
			Set objIntExplorer = Wscript.CreateObject("InternetExplorer.Application")			
		End IF
		If Not UCase(strPage) = "LOGGING" Then
			With objIntExplorer
				.Navigate strPage
				.ToolBar   = 0
				.Menubar   = 0
				.StatusBar = 1
			End With
			objIntExplorer.Visible = 1
			'Wait for IE to finish
			Do While (objIntExplorer.Busy)
				Wscript.Sleep 50
			Loop
		Else
			With objIntExplorer
				.Navigate "about:blank"
				.ToolBar   = 0
				.Menubar   = 0
				.StatusBar = 0
				.Width     = 710
				.Height    = 625
				.Left      = 100
				.Top       = 100
			End With
			'Set some formating
			With objIntExplorer.Document
				.WriteLn ("<!doctype html public>")
				.WriteLn   ("<head>")
				.WriteLn    ("<title>" & strTitle & "</title>")
				.WriteLn     ("<style type=""text/css"">")
				.WriteLn      ("body {text-align: left; font-family: verdana; font-size: 10pt}")
				.WriteLn     ("</style>")
				.WriteLn   ("</head>")
			End With
			'Show IE
			objIntExplorer.Visible = 1
			'Wait for IE to finish
			Do While (objIntExplorer.Busy)
				Wscript.Sleep 50
			Loop
		End If

		'Make IE the active window
		'For intCount = 1 To 100
			'If objWshShell.AppActivate(strTitle) Then Exit For
			'WScript.Sleep 50
		'Next

	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM UserPrompt Sub
	Private Sub UserPrompt( strPrompt )
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	UserPrompt
		' Purpose:  	Use Internet Explorer as a status message window
		' Input:    	strPrompt
		' Output:   	Output is sent to the open Internet Explorer window
		' Dependencies	objIntExplorer
		' Usage:
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	    'On Error Resume Next
	    objIntExplorer.Document.WriteLn (strPrompt & "<br />")
	End Sub
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Recursively creates all folders in a path
	Sub RCreateFolder( strPath )
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	RCreateFolder
		' Purpose:  	Make sure that only one Logon script is running per user
		' Input:		strPath
		' Output:
		' Dependencies	None
		' Usage:    	Call RCreateFolder [Folder and Folder path to create] 
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	If Not objFileSys.FolderExists( objFileSys.GetParentFolderName(strPath) ) then
			Call RCreateFolder( objFileSys.GetParentFolderName(strPath) )
		End If	
		objFileSys.CreateFolder( strPath )
	End Sub 
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
REM Single Instance
	Sub SingleInstance()
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	SingleInstance
		' Purpose:  	Make sure that only one Logon script is running per user
		' Input:
		' Output:
		' Dependencies	None
		' Usage:    	Call SingleInstance
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'On Error Resume Next
		Dim objWMIService
		Dim objWMIResults
		Dim objWMILoop
		Dim strWMIQuery
		Dim intCount
		intCount = 9999999999999
		
		Set objWMIService		  = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		' Uses WMI to see of the script is running and if so exits.
		strWMIQuery = "SELECT * FROM Win32_Process WHERE CommandLine LIKE '%" & wscript.scriptname & "%'"
		Set objWMIResults = objWMIService.ExecQuery(strWMIQuery)
		If objWMIResults.count > 1 Then 
			'Loop to find the newest script
			For Each objWMILoop in objWMIResults
				If intCount > objWMILoop.ProcessId Then
					intCount = objWMILoop.ProcessId
				End if 
			Next
			'Loop again to kill newest script
			For Each objWMILoop in objWMIResults
				If Not intCount = objWMILoop.ProcessId Then
					objWMILoop.Terminate()
					If Not Err.Number = 0 Then objWshShell.Popup Err.Number  & " Description: " & Err.Description & " Source: " & Err.Source
				End if 
			Next
		End If 
		
	End Sub
REM Cleanup Sub
	Sub Cleanup
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		' Sub:      	Cleanup
		' Purpose:  	Release common objects and exit script
		' Input:
		' Output:
		' Dependencies	All
		' Usage:    	Call Cleanup
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'On Error Resume Next
		'Set objUser 	   = Nothing
		Set objFileSys     = Nothing
		Set objWshNetwork  = Nothing
		Set objWshShell    = Nothing
		Set objIntExplorer = Nothing
		Set objNTUser 	   = Nothing

		'Exit script
		Wscript.Quit( )
	End Sub
