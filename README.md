<div align="center">

## Inventory\.vbs


</div>

### Description

This script is the first part of an inventory system. It uses WMI to read various information about a system. What makes it unique is that is opens an Internet Explorer application and creates a form containing the inventory data. It then automatically submits this form to an asp page sitting on my intranet which parses the data and loads in into an access db.
 
### More Info
 
I hope I have documented this code well enough for you to understand.

DEVELOPMENT HAS STOPPED on this project. I am moving and changing jobs. Sorry folks.

Belowe is the asp page and vbs script. Make sure you paste them into seperate files.

I do not have a good copy of the inventory DB. It contained proprietary data and I could not port it to the web. Sorry again.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bryan Beaty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bryan-beaty.md)
**Level**          |Advanced
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__4-7.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bryan-beaty-inventory-vbs__4-7259/archive/master.zip)

### API Declarations

This software is copyright (c) 2002 Sinton ISD. It is distributed under the GNU license. See comments for more details.


### Source Code

```
' ------------------------------------------------
' Inventory collection agent
' (C)2002 Sinton ISD
' Written by: Bryan Beaty
' This software is copywritten by Sinton ISD
' It may be distributed under the terms of the
' GNU General Public License at
' http://www.gnu.org/licenses/gpl.txt
' Any other use is prohibited.
'
' This script will collect PC information
' output it to an Internet Explorer Object
' and submit it to a web server via a web form.
' The asp page and database are required to make
' this software work.
' This script also uses WMI. WMI must be
' installed on the target machine.
' This is intended to follow a logon script.
'
' REVISION HISTORY
'
' Version 1.2.2
' Fixed: If a string value is blank then WMI could not
' 	retrieve it. It is either not specified by the manufacturer
' 	or OS. I have changed null or zero length strings to
' 	"Unknown" so they can be entered into the DB.
' 	This prevents errors on the asp page.
'	This should really be done on the asp side for security
'	reasons but it was faster to fix here.
' Version 1.2.1
' Added: If asp error occurs it asks the user
' 	to print and send the error to technology.
' Version 1.2
' Added: new tabels for each device. Each device becomes an
' 	object and is stored in its own table. This allows
' 	for much more data to be collected about each object.
' Added: use the PNP ID when possible to identify devices.
'	This allows for more accurate matching of products then
'	text matches allows for.
' Updated: Status messages now use dhtm so the stay on the screen.
' Known bugs: system hangs permanetly if wmi is not installed
' 	correctly.
'
' Version 1.1
' fixed: names with one digit room number will be parsed correctly.
' fixed: error with null parsing have been fixed
' added: status text to window so user doesn't get bored and we can
' 	track where errors occur.
' ------------------------------------------------
ScriptVersion="1.2.1"
' ------------------------------------------------
' Don't crash on error
' Remark out if you want to test script
' If you rem it out be aware that the regread
' function will generate an error if the key does
' not exist. This means that if this is the
' first time the script has run an error will be
' generated. This error is used to indicate that
' this is the first time it has run.
' ------------------------------------------------
ON ERROR RESUME NEXT
' ------------------------------------------------
' Set up vars
' ------------------------------------------------
Dim strDeviveType
Dim dateDevDate
Dim strName
Dim strSpeed
Dim intSize
Dim strDescription
Dim strComputer
Dim strNicMAC
Dim intRunInterval
Dim MSIE
' ------------------------------------------------
' Network Adapter Object
' AdapterType, AutoSense, Caption, Description, Manufacturer
' MaxSpeed, Name, PNPDeviceID, ProductName, ServiceName
' ------------------------------------------------
Dim objNA
' ------------------------------------------------
' Sound Card Object
' Caption, Description, DMABufferSize, Manufacturer, Name, PNPDeviceID, ProductName
' ------------------------------------------------
Dim objSC
' ------------------------------------------------
' Video Card Object
' AdapterCompatibility, AdapterRAM, Caption, Description, PNPDeviceID, VideoArchitecture
' VideoMemoryType, VideoProcessor
' ------------------------------------------------
Dim objVC
' ------------------------------------------------
' MotherBoard Object
' Caption, Description, Manufacturer, Model, Name, OtherIdentifyingInformation,
' PartNumber, Product, SKU, Version
' ------------------------------------------------
Dim objMB
' ------------------------------------------------
' Disk Drive Object
' Description, DeviceID, FileSystem, FreeSpace, Name, PNPDeviceID, Size
' ------------------------------------------------
Dim objDD
' ------------------------------------------------
' CD-ROM Object
' Description, Manufacturer, Name, PNPDeviceID
' ------------------------------------------------
Dim objCD
' ------------------------------------------------
' Determine if the program needs to run.
' intRunInterval is the number of days to wait before
' this program runs again. 0=run every time the
' system boots. 90 = run every 90 days.
' ------------------------------------------------
intRunInterval=1
' ------------------------------------------------
' Assume you need to run unless you test different.
' This reduces the number of else statements I need.
' bolRunFLag is used to indicate if the script needs
' to run
' ------------------------------------------------
bolRunFlag=True
set SHELL=CreateObject("WScript.Shell")
RegValue="HKEY_LOCAL_MACHINE\Software\SintonISD\InventoryRunDate"
If intRunInterval > 1 then
	dateLastRun=Shell.RegRead(regValue)
	if isDate(dateLastRun) then
		If DateDiff("d", dateLastRun, date()) < intRunInterval then
			bolRunFlag=False
		End IF
	End IF
End IF
If bolRunFlag=True then
	bolErrFlag=shell.regwrite(regvalue,date,"REG_SZ")
	Call DoInv
End If
Sub DoInv
	Call subOpenMSIE
	MSIE.Document.Write "<HTML>"
	MSIE.Document.Write "<HEAD><TITLE>PLEASE WAIT: Submitting Data.</TITLE></HEAD><BODY>" + vbCrLf
	MSIE.Document.Write "<CENTER><H1 ID='idHeader'>PLEASE WAIT</H1>"
	MSIE.Document.Write "<BR>"
	MSIE.Document.Write "<h3 ID='idStatus'>Gathering data</H3>"
	MSIE.Document.Write "<BR><BR>" + vbCrLf
	MSIE.Document.Write "<h3 ID='idInfo'></H3></CENTER>" + vbCrLf
	' ------------------------------------------------
	' Gather object information.
	' Each item below is put in a seperate table.
	' ------------------------------------------------
	Call subGetCDRom
	Call subGetDisk
	Call subGetMB
	call subGetVideo
	Call subGetSoundCard
	Call subGetNic
	' ------------------------------------------------
	' Gather System Name
	' ------------------------------------------------
	Set MBSet = GetObject("Winmgmts:").InstancesOf("Win32_MotherboardDevice")
	For Each MB in MBSet
		strComputerName=MB.SystemName
	NEXT
	IF len(strComputerName) > 7 THEN
		intPos=instr(strComputerName,"-")
		if intPos <> 0 then
			' if it has a dash it SHOULD be named correctly.
			strComputerRoom = right(strComputerName, len(strComputerName)-intPos)
			strComputerCampus = Left(strComputerName, 2)
			intComputerTag = Mid(strComputerName,3,intPos-3)
		END IF
	END IF
' ------------------------------------------------
' Gather Memory information
' ------------------------------------------------
inMemory=0
Set MemorySet = GetObject("Winmgmts:").InstancesOf("Win32_LogicalMemoryConfiguration")
For Each Memory in MemorySet
	intMemory=intMemory+Memory.TotalPhysicalMemory
NEXT
intMemory=round((intMemory/1024),0)
' ------------------------------------------------
' Gather OS Info
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "OS"
Set OSSet = GetObject("Winmgmts:").InstancesOf("Win32_OperatingSystem")
For Each OS in OSSet
	strOSname=OS.Caption
NEXT
' ------------------------------------------------
' Gather Processor information
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Processor"
Set ProSet = GetObject("Winmgmts:").InstancesOf("Win32_Processor")
For Each Pro in ProSet
	strProName=Pro.Name
	strProSpeed=Pro.CurrentClockSpeed
	strProManufacturer=Pro.Manufacturer
NEXT
' ------------------------------------------------
' Gather MAC Address
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "MAC"
If strNICMAC="" then
	Set NICSet = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
	For Each NIC in NICSet
		strNicMAC=NIC.MACaddress
	NEXT
End IF
' ------------------------------------------------
' Gather Printer Info
' ------------------------------------------------
'MSIE.document.all.idInfo.innerText = "Printer"
'Set PrinterSet = GetObject("Winmgmts:").InstancesOf("Win32_Printer")
'For Each Printer in PrinterSet
'	PrinterDesc=int(Printer.Attributes)
'	' This is a really funky way of finding out if it is a local printer.
'	' The returned value is the numeric representation of a binary value.
'	' I can't find any binary conversion tools in VBScript so I am
'	' doing it in base 10. There has to be an easier way but I am too
'	' fried to figure it out.
'	If PrinterDesc > 8091 then PrinterDesc=PrinterDesc-8092
'	If PrinterDesc > 4095 then PrinterDesc=PrinterDesc-4096
'	If PrinterDesc > 2047 then PrinterDesc=PrinterDesc-2048
'	If PrinterDesc > 1023 then PrinterDesc=PrinterDesc-1024
'	If PrinterDesc > 511 then PrinterDesc=PrinterDesc-512
'	If PrinterDesc > 255 then PrinterDesc=PrinterDesc-256
'	If PrinterDesc > 127 then PrinterDesc=PrinterDesc-128
'	If PrinterDesc > 63 then PrinterConnection = "Local"
'
'	Call subPrepDevicesObject
'	strName=Printer.DriverName
'	strDescription=PrinterConnection
'	Call subCleanDeviceObject
'	Call subAddDeviceObject
'
'	PrinterConnection = "Network"
'NEXT
' ------------------------------------------------
' Gather Keyboard information
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "KB"
intDeviceType=6
Set KbdSet = GetObject("Winmgmts:").InstancesOf("Win32_Keyboard")
For Each Kbd in KbdSet
	strKbd=kbd.Description
NEXT
' ------------------------------------------------
' Gather Mouse information
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Mouse"
intDeviceType=7
Set MouseSet = GetObject("Winmgmts:").InstancesOf("Win32_PointingDevice")
For Each Mouse in MouseSet
	strMouse=Mouse.HardwareType
NEXT
' ------------------------------------------------
' Gather Application Info
' ------------------------------------------------
'MSIE.document.all.idInfo.innerText = "Applications"
'intDeviceType=99
'Set AppSet = GetObject("Winmgmts:").InstancesOf("Win32_Product")
'For Each App in AppSet
'NEXT
' ------------------------------------------------
' Take all commas and apostrophes out of any data.
' Commas are my delimiter and will fry my asp routine
' if I don't delete them. They should be unneccesary anyway.
' Apostrophes will wreck my SQL statements in the ASP
' script. Don't need that either.
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Parsing Data"
if not isnull(strMBManufacturer) then strMBManufacturer=replace(strMBManufacturer,","," ")
if not isnull(strMBversion) then strMBversion=replace(strMBversion,","," ")
if not isnull(strMBtype) then strMBtype=replace(strMBtype,","," ")
if not isnull(strProName) then strProName=replace(strProName,","," ")
if not isnull(strProSpeed) then strProSpeed=replace(strProSpeed,","," ")
if not isnull(strProManufacturer) then strProManufacturer=replace(strProManufacturer,","," ")
if not isnull(strOSname) then strOSname=replace(strOSname,","," ")
if not isnull(strNICMAC) then strNICMAC=replace(strNICMAC,","," ")
if not isnull(strComputerCampus) then strComputerCampus=replace(strComputerCampus,","," ")
if not isnull(strComputerRoom) then strComputerRoom=replace(strComputerRoom,","," ")
if not isnull(strMBManufacturer) then strMBManufacturer=replace(strMBManufacturer,"'"," ")
if not isnull(strMBversion) then strMBversion=replace(strMBversion,"'"," ")
if not isnull(strMBtype) then strMBtype=replace(strMBtype,"'"," ")
if not isnull(strProName) then strProName=replace(strProName,"'"," ")
if not isnull(strProSpeed) then strProSpeed=replace(strProSpeed,"'"," ")
if not isnull(strProManufacturer) then strProManufacturer=replace(strProManufacturer,"'"," ")
if not isnull(strOSname) then strOSname=replace(strOSname,"'"," ")
if not isnull(strNICMAC) then strNICMAC=replace(strNICMAC,"'"," ")
if not isnull(strComputerCampus) then strComputerCampus=replace(strComputerCampus,"'"," ")
if not isnull(strComputerRoom) then strComputerRoom=replace(strComputerRoom,"'"," ")
' ------------------------------------------------
' Some values have a lot of white space.
' I am going to get rid of any whitespace in the str values.
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Trimming Whitespace"
strMBManufacturer=ltrim(rtrim(strMBManufacturer))
strMBversion=ltrim(rtrim(strMBversion))
strMBtype=ltrim(rtrim(strMBtype))
strProName=ltrim(rtrim(strProName))
strProSpeed=ltrim(rtrim(strProSpeed))
strProManufacturer=ltrim(rtrim(strProManufacturer))
strOSname=ltrim(rtrim(strOSname))
strNICMAC=ltrim(rtrim(strNICMAC))
strComputerCampus=ltrim(rtrim(strComputerCampus))
strComputerRoom=ltrim(rtrim(strComputerRoom))
' ------------------------------------------------
' I don't want too many invalid values
' If the values are not of the correct type
' I will blank them out in the correct type.
' If the date is invalid I am setting it to
' 1/1/1492. That should be a red flag when I cruise
' the data.
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Validating Data"
If strOSname = "" then strOSname = "Unknown"
If not isDate(dateBDate) then dateBdate="01/01/1492"
If not isNumeric(intMrmory) then intMemory=0
If not isNumeric(intComputerTag) then intComputerTag=0
' ------------------------------------------------
' Build output string to be parsed by the asp script
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "Building output strings: "
strComputer = strNicMAC & "," & intComputerTag & ","
strComputer = strComputer & strComputerCampus & "," & strComputerRoom & "," & intMemory & ","
strComputer = strComputer & strProManufacturer & "," & strProName & "," & strProSpeed & ","
strComputer = strComputer & strOSname & "," & date() & "," & strComputerName & ",##END##"
MSIE.document.all.idStatus.innerText = "SUBMITTING DATA"
' ------------------------------------------------
' Output form to MSIE object
' ------------------------------------------------
MSIE.document.all.idInfo.innerText = "If this window does not close automatically in 10 seconds click on the submit button below."
' ------------------------------------------------
' This script will submit the form when the page is
' completely loaded.
' ------------------------------------------------
MSIE.Document.Write "<SCRIPT TYPE=" & chr(34) & "text/vbscript" & chr(34) & ">" + vbCrLf
MSIE.Document.Write "Sub Window_onLoad" + vbCrLf
MSIE.Document.Write "	oForm.oSubmit.Click" + vbCrLf
MSIE.Document.Write "End Sub" + vbCrLf
MSIE.Document.Write "</SCRIPT>" + vbCrLf
' ------------------------------------------------
' This is the form.
' ------------------------------------------------
MSIE.Document.Write "<form ID='oForm' method='POST' name='Inventory' action='http://pirate.sintonisd.net/inventory.asp'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='COMPUTER' VALUE='" & strComputer & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objNA' VALUE='" & objNA & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objSC' VALUE='" & objSC & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objVC' VALUE='" & objVC & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objMB' VALUE='" & objMB & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objDD' VALUE='" & objDD & "'>" + vbCrLf
MSIE.Document.Write "<INPUT TYPE='HIDDEN' NAME='objCD' VALUE='" & objCD & "'>" + vbCrLf
MSIE.Document.Write "<p><input type='submit' id='oSubmit' value='Submit' name='Submit'></p></form><BR><BR>" + vbCrLf
' ------------------------------------------------
' Use these lines to output the hidden fields
' This is used for debugging.
' ------------------------------------------------
'MSIE.Document.Write strComputer & "<BR><BR>" + vbCrLf
'MSIE.Document.Write strDevice + vbCrLf
MSIE.Document.Write "</BODY></HTML>" + vbCrLf
MSIE.Document.Close
' ------------------------------------------------
	' ------------------------------------------------
	' Wait for the form to submit and kill the window
	' I am done.
	' ------------------------------------------------
	DO
	Loop While MSIE.Busy
	strText=MSIE.document.body.innertext
	if inStr(strText,"ASP error") <> 0 then
		strText="An error occured trying to inventory your computer."
		strText=strText & "Please click OK and then choose print on dialog box that will open."
		strText=strText & "Send the printout to Technology."
		strText=strText & "You may then close all annoying windows."
		msgBox strText, vbOkOnly, "OOPS!"
		bSuccess = MSIE.document.execCommand("Print")
	Else
		Call subCloseMSIE
	END IF
End Sub
' ------------------------------------------------
' This cleans out the values left behind by the
' last device.
' ------------------------------------------------
Sub subPrepDevicesObject
	strName=""
	strManufacturer=""
	strDescription=""
End Sub
' ------------------------------------------------
' This cleans the data so it doesn't fry the
' SQL statements or parser.
' ------------------------------------------------
Sub subCleanDeviceObject
	if not isnull(strName) then
		strName=ltrim(rtrim(strName))
	 	strName=replace(strName,"'"," ")
		strName=replace(strName,","," ")
	End If
	if not isnull(strDescription) then
		strDescription=ltrim(rtrim(strDescription))
		strDescription=replace(strDescription,"'"," ")
		strDescription=replace(strDescription,","," ")
	End If
	if not isnull(strManufacturer) then
		strManufacturer=ltrim(rtrim(strManufacturer))
		strManufacturer=replace(strManufacturer,"'"," ")
		strManufacturer=replace(strManufacturer,","," ")
	End IF
End Sub
' ------------------------------------------------
' This adds the values to a string that is
' comma delimited. Will be parsed by ASP script.
' This is why it was cleaned above.
' ------------------------------------------------
Sub subAddDeviceObject
	if len(strDevice) > 0 then strDevice=strDevice & ","
	strDevice = strDevice & intDeviceType & ","
	strDevice = strDevice & strName & ","
	strDevice = strDevice & strManufacturer & ","
	strDevice = strDevice & strDescription
End Sub
sub subCloseMSIE
	MSIE.Document.Close
	MSIE.quit
	set MSIE=nothing
End Sub
sub subOpenMSIE
	' ------------------------------------------------
	' Open Internet Explorer for writing: AGain
	' ------------------------------------------------
	Set MSIE=CreateObject("InternetExplorer.Application")
	' ------------------------------------------------
	' Set application settings so people can't mess with it.
	' ------------------------------------------------
	MSIE.Navigate "about:Blank"
	MSIE.Toolbar = False
	MSIE.StatusBar = False
	MSIE.Resizable = False
	' ------------------------------------------------
	' You may want to use these if you are debugging.
	' ------------------------------------------------
	'MSIE.Toolbar = True
	'MSIE.StatusBar = True
	'MSIE.Resizable = True
	' ------------------------------------------------
	' Wait for the app to be ready
	' ------------------------------------------------
	DO
	Loop While MSIE.Busy
	' ------------------------------------------------
	' size the page to fit half the screen
	' ------------------------------------------------
	SWidth = MSIE.Document.ParentWindow.Screen.AvailWidth
	SHeight = MSIE.Document.ParentWindow.Screen.AvailHeight
	MSIE.Width =SWidth/2
	MSIE.Height=SHeight/2
	MSIE.Left=(SWidth-MSIE.Width)/2
	MSIE.Top=(SHeight-MSIE.Height)/2
	MSIE.Visible = True
End Sub
Function fnCleanString(StringValue)
	If not isNull(StringValue) then
		StringValue=ltrim(rtrim(StringValue))
		StringValue=replace(StringValue,"'"," ")
		StringValue=replace(StringValue,","," ")
		if StringValue="" then StringValue="Unknown"
		fnCleanString=StringValue
	Else
		StringValue="Unknown"
	End IF
End Function
Function fnCleanNumber(NumberValue)
	If not isNumeric(NumberValue) then
		fnCleanNumber=0
	Else
		fnCleanNumber=StringValue
	End IF
End Function
sub subGetNic
	' ------------------------------------------------
	' Gather NIC information
	' ------------------------------------------------
	MSIE.document.all.idInfo.innerText = "NIC"
	Set NICSet = GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapter")
	objNA=""
	For Each NIC in NICSet
		if instr(NIC.PNPDeviceID,"VEN") <> 0 then
			objNA = objNA & fnCleanString(NIC.PNPDeviceID) & ","
			objNA = objNA & fnCleanString(NIC.AdapterType) & ","
			objNA = objNA & NIC.AutoSense & ","
			objNA = objNA & fnCleanString(NIC.Caption) & ","
			objNA = objNA & fnCleanString(NIC.Description) & ","
			objNA = objNA & fnCleanString(NIC.Manufacturer) & ","
			objNA = objNA & fnCleanNumber(NIC.MaxSpeed) & ","
			objNA = objNA & fnCleanString(NIC.Name) & ","
			objNA = objNA & fnCleanString(NIC.ProductName) & ","
			objNA = objNA & fnCleanString(NIC.ServiceName) & ","
			if NIC.MACAddress <> "" then
				strNicMAC=NIC.MACaddress
			End If
		End If
	NEXT
	objNA = objNA & "##END##"
End Sub
sub subGetSoundCard
	' ------------------------------------------------
	' Gather Sound Card information
	' PNPDeviceID, Caption, Description, DMABufferSize, Manufacturer, Name, ProductName
	' ------------------------------------------------
	MSIE.document.all.idInfo.innerText = "Sound Card"
	Set SoundSet = GetObject("Winmgmts:").InstancesOf("Win32_SoundDevice")
	objSC=""
	For Each Sound in SoundSet
		 if instr(Sound.PNPDeviceID,"VEN") <> 0 then
			objSC = objSC & fnCleanString(Sound.PNPDeviceID) & ","
			objSC = objSC & fnCleanString(Sound.Caption) & ","
			objSC = objSC & fnCleanString(Sound.Description) & ","
			objSC = objSC & fnCleanNumber(Sound.DMABufferSize) & ","
			objSC = objSC & fnCleanString(Sound.Manufacturer) & ","
			objSC = objSC & fnCleanString(Sound.Name) & ","
			objSC = objSC & fnCleanString(Sound.ProductName) & ","
		End If
	NEXT
	objSC = objSC & "##END##"
End Sub
sub subGetVideo
	' ------------------------------------------------
	' Gather Video Card information
	' PNPDeviceID, AdapterCompatibility, AdapterRAM, Caption, Description, VideoArchitecture
	' VideoMemoryType, VideoProcessor
	' ------------------------------------------------
	MSIE.document.all.idInfo.innerText = "Video Card"
	objVC=""
	Set VideoSet = GetObject("Winmgmts:").InstancesOf("Win32_VideoController")
	For Each Video in VideoSet
		objVC = objVC & fnCleanString(Video.PNPDeviceID) & ","
		objVC = objVC & fnCleanString(Video.AdapterCompatibility) & ","
		objVC = objVC & fnCleanNumber(Video.AdapterRAM) & ","
		objVC = objVC & fnCleanString(Video.Caption) & ","
		objVC = objVC & fnCleanString(Video.Description) & ","
		objVC = objVC & fnCleanNumber(Video.VideoArchitecture) & ","
		objVC = objVC & fnCleanNumber(Video.VideoMemoryType) & ","
		objVC = objVC & fnCleanString(Video.VideoProcessor) & ","
	NEXT
	objVC = objVC & "##END##"
End Sub
sub subGetMB
	' ------------------------------------------------
	' Gather BaseBoard information
	' Caption, Description, Manufacturer, Model, Name, OtherIdentifyingInfo,
	' PartNumber, Product, SKU, Version
	' ------------------------------------------------
	objMB=""
	MSIE.document.all.idInfo.innerText = "Baseboard"
	Set MBSet = GetObject("Winmgmts:").InstancesOf("Win32_BaseBoard")
	For Each MB in MBSet
		objMB = objMB & fnCleanString(MB.Product) & ","
		objMB = objMB & fnCleanString(MB.Caption) & ","
		objMB = objMB & fnCleanString(MB.Description) & ","
		objMB = objMB & fnCleanString(MB.Manufacturer) & ","
		objMB = objMB & fnCleanString(MB.Name) & ","
		objMB = objMB & fnCleanString(MB.OtherIdentifyingInfo) & ","
		objMB = objMB & fnCleanString(MB.PartNumber) & ","
		objMB = objMB & fnCleanString(MB.Model) & ","
		objMB = objMB & fnCleanString(MB.SKU) & ","
		objMB = objMB & fnCleanString(MB.Version) & ","
	NEXT
	objMB = objMB & "##END##"
End Sub
Sub subGetDisk
	' ------------------------------------------------
	' Gather Disk information
	' Description, DeviceID, FileSystem, FreeSpace, Name, PNPDeviceID, Size
	' ------------------------------------------------
	MSIE.document.all.idInfo.innerText = "Disk"
	Set DiskSet = GetObject("Winmgmts:").InstancesOf("Win32_LogicalDisk")
	objDD=""
	For Each Disk in DiskSet
		Select Case Disk.DriveType
			Case 3
				objDD = objDD & fnCleanString(Disk.Description) & ","
				objDD = objDD & fnCleanString(Disk.DeviceID) & ","
				objDD = objDD & fnCleanString(Disk.FileSystem) & ","
				objDD = objDD & fnCleanString(Disk.FreeSpace) & ","
				objDD = objDD & fnCleanString(Disk.Name) & ","
				objDD = objDD & fnCleanString(Disk.PNPDeviceID) & ","
				objDD = objDD & fnCleanString(Disk.Size) & ","
		End Select
	NEXT
	objDD = objDD & "##END##"
End Sub
Sub subGetCDRom
	' ------------------------------------------------
	' Gather CD-ROM information
	' Description, Manufacturer, Name, PNPDeviceID
	' ------------------------------------------------
	MSIE.document.all.idInfo.innerText = "CD"
	objCD=""
	Set CDSet = GetObject("Winmgmts:").InstancesOf("Win32_CDROMDrive")
	For Each CD in CDSet
		objCD= objCD & fnCleanString(CD.Description) & "Description,"
		objCD= objCD & fnCleanString(CD.Manufacturer) & "Manufacturer,"
		objCD= objCD & fnCleanString(CD.Name) & "Name,"
		objCD= objCD & fnCleanString(CD.PNPDeviceID) & "PNPDeviceID,"
	NEXT
	objDD = objDD & "##END##"
End Sub
<--- end vbs file
****************************************
****************************************
****************************************
Here is the asp page:
***************************************
***************************************
***************************************
start asp page -->
<%
	' -------------------------------------------------------
	' Inventory.asp
	' (c)2002 Sinton ISD
	' Witten by: Bryan Beaty
	' This may be distrributed under the GNU public license
	'
	' Version 1.2 beta
	' -------------------------------------------------------
	Response.Expires=0
	' -------------------------------------------------------
	' Define vars used throughout the program
	' -------------------------------------------------------
	Dim objConn, objRS, strQuery, strConnection
	' -------------------------------------------------------
	' These vars hold the comma delimited data submitted by the form.
	' -------------------------------------------------------
	Dim strComputer, strNA, strSC, strVC, strMB, strDD, strCD
	' -------------------------------------------------------
	' These vars hold other critical data
	' -------------------------------------------------------
	Dim intComputerType
	' -------------------------------------------------------
	' These arrays hold the parsed data from the above strings.
	' The array size is from zero so there are 7 fields if if it is arr(6)
	' -------------------------------------------------------
	Dim arrComputer(10), arrNA(9), arrSC(6), arrVC(7)
	Dim arrMB(9), arrDD(8), arrCD(3)
	' -------------------------------------------------------
	' This is used to parse data. It needs to be as big as the
	' largest array above.
	' -------------------------------------------------------
	Dim arrTEMP(10)
	Dim strTEMP
	' -------------------------------------------------------
	' These vars are for various routines
	' -------------------------------------------------------
	Dim bolError, strDebug
	Dim i
	Dim intDeviceID
	' -------------------------------------------------------
	' Pull the data submitted from the server
	' -------------------------------------------------------
	strComputer=trim(Request.Form("Computer"))
	strNA=trim(Request.Form("objNA"))
	strSC=trim(Request.Form("objSC"))
	strVC=trim(Request.Form("objVC"))
	strMB=trim(Request.Form("objMB"))
	strDD=trim(Request.Form("objDD"))
	strCD=trim(Request.Form("objCD"))
	' -------------------------------------------------------
	' strXX's are comma delimeted strings
	' I will parse them into the arrTEMP array and
	' update the appropriate arrays.
	' -------------------------------------------------------
	Call subOpenDB
	Call subGetMB
	Call subGetVC
	Call subGetSC
	Call subGetNA
	Call subGetComputerType
	Call subGetComputer
	Call subCloseDB
Sub subOpenDB
	Set objConn=Server.CreateObject("ADODB.Connection")
	strConnection = "DSN=WR2001;Database=WR2001;"
	objConn.Open strConnection
End Sub
sub subGetMB
	do until strMB="##END##"
		bolError=fnPopArray(10, strMB)
		strMB=strTEMP
		for i = 0 to 9
			arrMB(i)=arrTEMP(i)
		Next
		If arrMB(0)="" then
			If arrMB(8)="" then
				arrMB(0)="Unknown"
			Else
				arrMB(0)=arrMB(8)
			End If
		End IF
		If bolError=0 then
			Call subUpdateTblMB
		End IF
	Loop
End Sub
Function fnPopArray(intFieldCount, strData)
	' -------------------------------------------------------
	' I should have intFieldCount items in this list.
	' -------------------------------------------------------
	strTemp=strData
	For i = 0 to (intFieldCount-1)
		intPos=instr(strTemp,",")
		if intPos <> 0 then
			If intPos=1 then
				arrTEMP(i)=""
			Else
				arrTEMP(i)=left(strTEMP,intPos-1)
				' If the length is greater than 255 chars then crop it.
				if len(arrTEMP(i))> 255 then arrTEMP(i)=left(arrTEMP(i),255)
			End If
			strTEMP=right(strTEMP,len(strTEMP)-intPos)
		Else
			arrTEMP(i)=strTEMP
		End If
	Next
	fnPopArray=0
End Function
Sub subUpdateTblMB
	strQuery = "SELECT Count(*) as intCountOfDevices FROM tblMB WHERE Product='" & arrMB(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("intCountOfDevices") = 0 then
		If arrMB(0)="" then
			If arrMB(8)="" then
				arrMB(0)="Unknown"
			Else
				arrMB(0)=arrMB(8)
			End If
		End IF
		' Add a new record to the tblMB table
		strQuery = "INSERT INTO tblMB (Product, Caption, Description, Manufacturer, "
		strQuery = strQuery & "Name, OtherIdentifyingInfo, PartNumber, Model, SKU, Version) VALUES "
		strQuery = strQuery & "('" & arrMB(0) & "',"
		strQuery = strQuery & "'" & arrMB(1) & "',"
		strQuery = strQuery & "'" & arrMB(2) & "',"
		strQuery = strQuery & "'" & arrMB(3) & "',"
		strQuery = strQuery & "'" & arrMB(4) & "',"
		strQuery = strQuery & "'" & arrMB(5) & "',"
		strQuery = strQuery & "'" & arrMB(6) & "',"
		strQuery = strQuery & "'" & arrMB(7) & "',"
		strQuery = strQuery & "'" & arrMB(8) & "',"
		strQuery = strQuery & "'" & arrMB(9) & "')"
	ELSE
		' Update an old record in the DB
		' Model is the only thing that stays the same.
		' Update everything else.
		strQuery = "UPDATE tblMB SET "
		strQuery = strQuery & "Caption='" & arrMB(1) & "',"
		strQuery = strQuery & "Description='" & arrMB(2) & "',"
		strQuery = strQuery & "Manufacturer='" & arrMB(3) & "',"
		strQuery = strQuery & "Name='" & arrMB(4) & "',"
		strQuery = strQuery & "OtherIdentifyingInfo='" & arrMB(5) & "',"
		strQuery = strQuery & "PartNumber='" & arrMB(6) & "',"
		strQuery = strQuery & "Model='" & arrMB(7) & "',"
		strQuery = strQuery & "SKU='" & arrMB(8) & "',"
		strQuery = strQuery & "Version='" & arrMB(9) & "' "
		strQuery = strQuery & "WHERE Product='" & arrMB(0) & "';"
	END IF
	strDebug=strQuery
	set objRS=objConn.Execute(strQuery)
End Sub
sub subGetComputerType
	intPOS=instr(arrVC(0), "SUBSYS")
	if intPOS <> 0 then
		fnCleanPNP=left(arrVC(0),intpos-2)
	End IF
	intPOS=instr(arrMB(0), "SUBSYS")
	if intPOS <> 0 then
		fnCleanPNP=left(arrMB(0),intpos-2)
	End IF
	intPOS=instr(arrNA(0), "SUBSYS")
	if intPOS <> 0 then
		fnCleanPNP=left(arrNA(0),intpos-2)
	End IF
	strQuery = "SELECT Count(*) as CountOfDevices FROM tblComputerType WHERE NAID='" & arrNA(0) & "' AND "
	strQuery = strQuery & "VCID='" & arrVC(0) & "' AND "
	strQuery = strQuery & "MBID='" & arrMB(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("CountofDevices") = 0 then
	strQuery = "INSERT INTO tblComputerType (NAID, VCID, MBID) VALUES "
		strQuery = strQuery & "('" & arrNA(0) & "',"
		strQuery = strQuery & "'" & arrVC(0) & "',"
		strQuery = strQuery & "'" & arrMB(0) & "')"
		set objRS=objConn.Execute(strQuery)
	End If
	strQuery = "SELECT ID FROM tblComputerType WHERE NAID='" & arrNA(0) & "' AND "
	strQuery = strQuery & "VCID='" & arrVC(0) & "' AND "
	strQuery = strQuery & "MBID='" & arrMB(0) & "';"
	set objRS=objConn.Execute(strQuery)
	intComputerType=objRS("ID")
End Sub
sub subGetVC
	do until strVC="##END##"
		bolError=fnPopArray(8, strVC)
		strVC=strTEMP
		for i = 0 to 7
			arrVC(i)=arrTEMP(i)
			arrTEMP(i)=""
		Next
		If bolError=0 then
			if arrVC(2)="" then arrVC(2)=0
			if arrVC(5)="" then arrVC(5)=0
			if arrVC(6)="" then arrVC(6)=0
			Call subUpdateTblVC
		End IF
	Loop
End Sub
sub subGetSC
	do until strSC="##END##"
		bolError=fnPopArray(7, strSC)
		strSC=strTEMP
		for i = 0 to 6
			arrSC(i)=arrTEMP(i)
			arrTEMP(i)=""
		Next
		If bolError=0 then
			if arrSC(3)="" then arrSC(3)=0
			Call subUpdateTblSC
		End IF
	Loop
End Sub
sub subGetNA
	do until strNA="##END##"
		bolError=fnPopArray(10, strNA)
		strNA=strTEMP
		for i = 0 to 9
			arrNA(i)=arrTEMP(i)
			arrTEMP(i)=""
		Next
		If bolError=0 then
			if arrNA(2)="" then arrNA(2)=0
			if arrNA(6)="" then arrNA(6)=0
			Call subUpdateTblNA
		End IF
	Loop
end sub
sub subGetComputer
	bolError=fnPopArray(11, strComputer)
	for i = 0 to 10
		arrComputer(i)=arrTEMP(i)
		arrTEMP(i)=""
	Next
	If bolError=0 then
		if arrComputer(1)="" then arrComputer(1)=0
		if arrComputer(4)="" then arrComputer(4)=0
		Call subUpdateTblComputer
	End IF
End sub
Sub subUpdateTblVC
	If arrVC(0)="" then
		If arrVC(4)="" then
			If arrVC(3)="" then
				arrVC(0)="Unknown"
			Else
				arrVC(0)=arrVC(3)
			End If
		Else
			arrVC(0)=arrVC(4)
		End If
	End IF
	arrVC(0)=fnCleanPNP(arrVC(0))
	strQuery = "SELECT Count(*) as intCountOfDevices FROM tblVC WHERE PNPDeviceID='" & arrVC(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("intCountOfDevices") = 0 then
		' Add a new record to the tblVC table
		strQuery = "INSERT INTO tblVC (PNPDeviceID, AdapterCompatibility, AdapterRAM, Caption, "
		strQuery = strQuery & "Description, VideoArchitecture, VideoMemoryType, VideoProcessor) VALUES "
		strQuery = strQuery & "('" & arrVC(0) & "',"
		strQuery = strQuery & "'" & arrVC(1) & "',"
		strQuery = strQuery & arrVC(2) & ","
		strQuery = strQuery & "'" & arrVC(3) & "',"
		strQuery = strQuery & "'" & arrVC(4) & "',"
		strQuery = strQuery & arrVC(5) & ","
		strQuery = strQuery & arrVC(6) & ","
		strQuery = strQuery & "'" & arrVC(7) & "')"
	ELSE
		' Update an old record in the DB
		' PNPDeviceID is the only thing that stays the same.
		' Update everything else.
		strQuery = "UPDATE tblVC SET "
		strQuery = strQuery & "AdapterCompatibility='" & arrVC(1) & "',"
		strQuery = strQuery & "AdapterRAM='" & arrVC(2) & "',"
		strQuery = strQuery & "Caption='" & arrVC(3) & "',"
		strQuery = strQuery & "Description='" & arrVC(4) & "',"
		strQuery = strQuery & "VideoArchitecture=" & arrVC(5) & ","
		strQuery = strQuery & "VideoMemoryType=" & arrVC(6) & ", "
		strQuery = strQuery & "VideoProcessor='" & arrVC(7) & "' "
		strQuery = strQuery & "WHERE PNPDeviceID='" & arrVC(0) & "';"
	END IF
	strDebug=strQuery
	set objRS=objConn.Execute(strQuery)
End Sub
Sub subUpdateTblSC
	If arrSC(0)="" then
		If arrSC(5)="" then
			If arrSC(6)="" then
				arrSC(0)="Unknown"
			Else
				arrSC(0)=arrSC(6)
			End If
		Else
			arrSC(0)=arrSC(5)
		End If
	End IF
	arrSC(0)=fnCleanPNP(arrSC(0))
	strQuery = "SELECT Count(*) as intCountOfDevices FROM tblSC WHERE PNPDeviceID='" & arrSC(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("intCountOfDevices") = 0 then
		' Add a new record to the tblSC table
		strQuery = "INSERT INTO tblSC (PNPDeviceID, Caption, Description, DMABufferSize, "
		strQuery = strQuery & "Manufacturer, Name, ProductName) VALUES "
		strQuery = strQuery & "('" & arrSC(0) & "',"
		strQuery = strQuery & "'" & arrSC(1) & "',"
		strQuery = strQuery & "'" & arrSC(2) & "',"
		strQuery = strQuery & arrSC(3) & ","
		strQuery = strQuery & "'" & arrSC(4) & "',"
		strQuery = strQuery & "'" & arrSC(5) & "',"
		strQuery = strQuery & "'" & arrSC(6) & "')"
	ELSE
		' Update an old record in the DB
		' PNPDeviceID is the only thing that stays the same.
		' Update everything else.
		strQuery = "UPDATE tblSC SET "
		strQuery = strQuery & "Caption='" & arrSC(1) & "',"
		strQuery = strQuery & "Description='" & arrSC(2) & "',"
		strQuery = strQuery & "DMABufferSize=" & arrSC(3) & ","
		strQuery = strQuery & "Manufacturer='" & arrSC(4) & "',"
		strQuery = strQuery & "Name='" & arrSC(5) & "',"
		strQuery = strQuery & "ProductName='" & arrSC(6) & "' "
		strQuery = strQuery & "WHERE PNPDeviceID='" & arrSC(0) & "';"
	END IF
	strDebug=strQuery
	set objRS=objConn.Execute(strQuery)
End Sub
Sub subUpdateTblNA
	If arrNA(0)="" then
		If arrNA(7)="" then
			If arrNA(8)="" then
				arrNA(0)="Unknown"
			Else
				arrNA(0)=arrNA(8)
			End If
		Else
			arrNA(0)=arrNA(7)
		End If
	End IF
	arrNA(0)=fnCleanPNP(arrNA(0))
	strQuery = "SELECT Count(*) as intCountOfDevices FROM tblNA WHERE PNPDeviceID='" & arrNA(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("intCountOfDevices") = 0 then
		' Add a new record to the tblNA table
		strQuery = "INSERT INTO tblNA (PNPDeviceID, AdapterType, AutoSense, Caption, "
		strQuery = strQuery & "Description, Manufacturer, MaxSpeed, Name, ProductName, "
		strQuery = strQuery & "ServiceName) VALUES "
		strQuery = strQuery & "('" & arrNA(0) & "',"
		strQuery = strQuery & "'" & arrNA(1) & "',"
		strQuery = strQuery & arrNA(2) & ","
		strQuery = strQuery & "'" & arrNA(3) & "',"
		strQuery = strQuery & "'" & arrNA(4) & "',"
		strQuery = strQuery & "'" & arrNA(5) & "',"
		strQuery = strQuery & arrNA(6) & ","
		strQuery = strQuery & "'" & arrNA(7) & "',"
		strQuery = strQuery & "'" & arrNA(8) & "',"
		strQuery = strQuery & "'" & arrNA(9) & "')"
	ELSE
		' Update an old record in the DB
		' PNPDeviceID is the only thing that stays the same.
		' Update everything else.
		strQuery = "UPDATE tblNA SET "
		strQuery = strQuery & "AdapterType='" & arrNA(1) & "',"
		strQuery = strQuery & "AutoSense=" & arrNA(2) & ","
		strQuery = strQuery & "Caption='" & arrNA(3) & "',"
		strQuery = strQuery & "Description='" & arrNA(4) & "',"
		strQuery = strQuery & "Manufacturer='" & arrNA(5) & "',"
		strQuery = strQuery & "MaxSpeed=" & arrNA(6) & ","
		strQuery = strQuery & "Name='" & arrNA(7) & "',"
		strQuery = strQuery & "ProductName='" & arrNA(8) & "',"
		strQuery = strQuery & "ServiceName='" & arrNA(9) & "' "
		strQuery = strQuery & "WHERE PNPDeviceID='" & arrNA(0) & "';"
	END IF
	strDebug=strQuery
	set objRS=objConn.Execute(strQuery)
End Sub
Sub subUpdateTblComputer
	strQuery = "SELECT Count(*) as intCountOfDevices FROM tblComputers WHERE MAC='" & arrComputer(0) & "';"
	set objRS=objConn.Execute(strQuery)
	If objRS("intCountOfDevices") = 0 then
		' Add a new record to the tblCompter table
		strQuery = "INSERT INTO tblComputers (MAC, Tag, Campus, Room, TotalMemory, CPUManufacturer, CPUModel, CPUSpeed, OSType, DateInventoried, Name, ComputerType) VALUES "
		strQuery = strQuery & "('" & arrComputer(0) & "',"
		strQuery = strQuery & arrComputer(1) & ","
		strQuery = strQuery & "'" & arrComputer(2) & "',"
		strQuery = strQuery & "'" & arrComputer(3) & "',"
		strQuery = strQuery & arrComputer(4) & ","
		strQuery = strQuery & "'" & arrComputer(5) & "',"
		strQuery = strQuery & "'" & arrComputer(6) & "',"
		strQuery = strQuery & "'" & arrComputer(7) & "',"
		strQuery = strQuery & "'" & arrComputer(8) & "',"
		strQuery = strQuery & "#" & arrComputer(9) & "#,"
		strQuery = strQuery & "'" & arrComputer(10) & "',"
		strQuery = strQuery & "'" & intComputerType & "')"
	ELSE
		' Update an old record in the DB
		' MAC is the only thing that stays the same.
		' Update everything else.
		strQuery = "UPDATE tblComputers SET "
		strQuery = strQuery & "Tag=" & arrComputer(1) & ","
		strQuery = strQuery & "Campus='" & arrComputer(2) & "',"
		strQuery = strQuery & "Room='" & arrComputer(3) & "',"
		strQuery = strQuery & "TotalMemory=" & arrComputer(4) & ","
		strQuery = strQuery & "CPUManufacturer='" & arrComputer(5) & "',"
		strQuery = strQuery & "CPUModel='" & arrComputer(6) & "',"
		strQuery = strQuery & "CPUSpeed='" & arrComputer(7) & "',"
		strQuery = strQuery & "OSType='" & arrComputer(8) & "',"
		strQuery = strQuery & "DateInventoried=#" & arrComputer(9) & "#,"
		strQuery = strQuery & "Name='" & arrComputer(10) & "',"
		strQuery = strQuery & "ComputerType=" & intComputerType & " "
		strQuery = strQuery & "WHERE MAC='" & arrComputer(0) & "';"
	END IF
	set objRS=objConn.Execute(strQuery)
End Sub
Sub subCloseDB
	' Close connections
	objConn.close
	set objConn = nothing
End Sub
Function fnCleanPNP(strClean)
	intPOS=instr(strClean, "SUBSYS")
	if intPOS <> 0 then
		fnCleanPNP=left(strClean,intpos-2)
	End IF
End Function
%>
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inventory</title>
</head>
<body>
<H1>Completed</H1><BR>
You may close this page now.
Debug Info:
<%=strMB%><BR>
<%=strVC%><BR>
<%=strNA%><BR>
<%=strSC%><BR>
</body>
</html>
```

