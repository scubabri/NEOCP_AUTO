Set objShell = CreateObject("WScript.Shell")

' set your working directory here
baseDir = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"  

' set your minimums here
minscore = 90
mindec = 0 
minvmag = 20	
minobs = 4
minseen = .8

Sub Include(file)
	On Error Resume Next
	Dim FSO
	Set FSO = CreateObject("Scripting.FileSystemObject")
	ExecuteGlobal FSO.OpenTextFile(file & ".vbs", 1).ReadAll()
	Set FSO = Nothing

	If Err.Number <> 0 Then
		If Err.Number = 1041 Then
			Err.Clear
		Else
			WScript.Quit 1
		End If
	End If
End Sub

Include "VbsJson"																		' include json parser

Dim json, neocpStr, jsonDecoded
Set json = New VbsJson																	' define json parser

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO

scoutLink = "https://ssd-api.jpl.nasa.gov/scout.api?tdes="
scoutSaveFile = "\scout.json"
 
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
neocpFile = baseDir+"\neocp.txt"														' where to put the downloaded neocp.txt, adjust as required.
objectsSaveFile = baseDir+"\output.txt"													' where to put output of selected NEOCP objects for further parsing

orbLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
orbSaveFile = "\orbits.txt"

mpcorbSaveFile = baseDir+"\MPCORB.dat"													' the final (almost) MPCORB.dat 
 																				
objShell.CurrentDirectory = baseDir														' WGet saves file always on the actual folder. So, change the actual 
																						'folder for C:\, where we want to save file
Function Quotes(strQuotes)																' Add Quotes to string
	Quotes = chr(34) & strQuotes & chr(34)												' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
End Function
 
if objFSO.FileExists(mpcorbSaveFile) then												' get rid of stragglers 
	objFSO.DeleteFile mpcorbSaveFile
end if

if objFSO.FileExists(objectsSaveFile) then
	objFSO.DeleteFile objectsSaveFile
end if

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(neocpLink) & " -N",1,True 	'download current neocp.txt from MPC 
 
Set neocpFileRead = objFSO.OpenTextFile(neocpFile, 1) 									' change path for input file from wget 

Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(objectsSaveFile,8,true)  	' create output.txt
Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  	' MPCORB.dat output
		
Do Until neocpFileRead.AtEndOfStream
    strLine = neocpFileRead.ReadLine													' its probably a good idea NOT to touch the positions as they are fixed position.
	object = Mid(strLine, 1,7)															' temporary object designation
	score = Mid(strLine, 9,3)															' neocp desirablility score from 0 to 100, 100 being most desirable.
	ra	  = Mid(strLine, 27,7)
	dec = Mid(strLine, 35,6)															' declination 
	vmag = Mid(strLine, 44,4)															' if you dont know what this is, change hobbies
	obs = Mid(strLine, 79,4)															' how many observations has it had
	seen = Mid(strLine, 96,7)															' when was the object last seen
	
    if (CSng(score) >= minscore) AND (CSng(dec) >= mindec) AND (CSng(vmag) <= minvmag) AND (CSng(obs) >= minobs) AND (CSng(seen) <= minseen) Then
		
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(baseDir) & orbSaveFile,1,True ' run wget to get orbits from NEOCP
		
		Set objRegEx = CreateObject("VBScript.RegExp")									' This section parses the orbits from the neocp object and gets the NEOCPNomin parameters
		objRegEx.Pattern = "NEOCPNomin"													' which is used in the MPCORB.dat 
        Set objFile = objFSO.OpenTextFile(baseDir+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			
			If colMatches.Count > 0 Then
					MPCorbFileToWrite.WriteLine(strSearchString+"           "+object)	'write elemets to MPCORB.dat
				Exit Do
			End If
		Loop

		objFile.Close
		objectsFileToWrite.WriteLine(object+ "     " + score + "  " + ra + "  " + dec + "    " + vmag + "      " + obs + "     " + seen)	
	End If	
Loop

Set neocpFileRead = objFSO.OpenTextFile(neocpFile, 1) 

Do Until neocpFileRead.AtEndOfStream
    strLine = neocpFileRead.ReadLine													' its probably a good idea NOT to touch the positions as they are fixed position.
	object = Mid(strLine, 1,7)															' temporary object designation
	score = Mid(strLine, 9,3)															' neocp desirablility score from 0 to 100, 100 being most desirable.
	ra	  = Mid(strLine, 27,7)
	dec = Mid(strLine, 35,6)															' declination 
	vmag = Mid(strLine, 44,4)															' if you dont know what this is, change hobbies
	obs = Mid(strLine, 79,4)															' how many observations has it had
	seen = Mid(strLine, 96,7)															' when was the object last seen
	
	
    if (CSng(score) >= minscore) AND (CSng(dec) >= mindec) AND (CSng(vmag) <= minvmag) AND (CSng(obs) >= minobs) AND (CSng(seen) <= minseen) Then	
	   
	   objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(baseDir) & orbSaveFile,1,True ' run wget to get orbits from NEOCP
		
		Set objRegEx = CreateObject("VBScript.RegExp")									' This section parses the orbits from the neocp object and gets the NEOCPNomin parameters
		objRegEx.Pattern = "NEOCPNomin"													' which is used in the MPCORB.dat 
        Set objFile = objFSO.OpenTextFile(baseDir+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			
			If colMatches.Count > 0 Then
					mpcLine = strSearchString+"           "+object	'write elemets to MPCORB.dat
				Exit Do
			End If
		Loop
		
	   call GetExposureData(expTime,imageCount)
		
		
		If imageCount > 0 Then
		
		
			Dim RTML, REQ, TGT, PIC, COR, FSO, FIL                      ' (for PrimalScript IntelliSense)

			Set RTML = CreateObject("DC3.RTML23.RTML")
			Set RTML.Contact = CreateObject("DC3.RTML23.Contact")
			RTML.Contact.User = "Brian Sheets"
			RTML.Contact.Email = "brians@fl240.com"

			Set REQ =  CreateObject("DC3.RTML23.Request")
			REQ.UserName = "brians"                                  
			REQ.Project = "NEOCP"                                    ' Proj for above user will be created if needed

			Set REQ.Schedule = CreateObject("DC3.RTML23.Schedule")
			REQ.Schedule.Horizon = 30
			Set REQ.Schedule.Moon = CreateObject("DC3.RTML23.Moon")
			Set REQ.Schedule.Moon.Lorentzian = CreateObject("DC3.RTML23.Lorentzian")
			REQ.Schedule.Moon.Lorentzian.Distance  = 15
			REQ.Schedule.Moon.Lorentzian.Width = 6

			Set REQ.Correction = CreateObject("DC3.RTML23.Correction")
			REQ.Correction.zero = False
			REQ.Correction.flat = False
			REQ.Correction.dark = False
		
			RTML.RequestsC.Add REQ
			REQ.ID = object                                         ' This becomes the Plan name for the Request
			Set TGT = CreateObject("DC3.RTML23.Target")
			TGT.TargetType.OrbitalElements = mpcLine
			TGT.Name = object
			'TGT.count = 2
			'TGT.Interval=0.2  ' HOURS!!!
			REQ.TargetsC.Add TGT
		
			' An RTML Picture (child of Target) You may set the count property to acquire
			' multiple images within this Picture, and you may set the autostack property
			' to True to cause the repeated images to be automatically aligned and stacked.
			'
			Set PIC = CreateObject("DC3.RTML23.Picture")
			PIC.Name = object+" Clear"                               ' Required
			PIC.ExposureSpec.ExposureTime = expTime
			PIC.Binning = 2
			PIC.Filter = "Clear"
			PIC.Count = imageCount

			' Now add this Picture to the Target. You can repeat the above,
			' creating additional Pictures, and add them to the Target.
			' NOTE USE OF COM ACCESSOR.

			TGT.PicturesC.Add PIC
		
			XML = RTML.XML(True)
		
			Set FSO = CreateObject("Scripting.FileSystemObject")
			Set FIL = FSO.CreateTextFile("E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\NEOCP.rtml", True)    ' **CHANGE FOR YOUR SYSTEM**
			FIL.Write XML                                           ' Has embedded line endings
			FIL.Close
		
			Dim I, DB, R
			Set DB = CreateObject("DC3.Scheduler.Database")
			Call DB.Connect()
			Set I = CreateObject("DC3.RTML23.Importer")
			Set I.DB = DB
			I.Import "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\NEOCP.rtml"
			Set R = I.Projects.Item(0)
			R.Disabled = false
			R.Update()
            I.Plans.Item(0).Resubmit()
			Call DB.Disconnect()
			Set REQ =  nothing
			Set RTML = nothing
			Set TGT = nothing
			set PIC = nothing
		End If
	End If	
Loop

objFile.Close
objShell.Run "taskkill.exe /IM TheSkyX.exe"
		
neocpFileRead.Close																		' close any open files
objectsFileToWrite.Close
MPCorbFileToWrite.Close

if objFSO.FileExists(baseDir+orbSaveFile) then											' clean up temporary files
	objFSO.DeleteFile basedir+orbSaveFile
end if
if objFSO.FileExists(neocpFile) then
	objFSO.DeleteFile neocpFile
end if
if objFSO.FileExists(baseDir+scoutSaveFile) then
	objFSO.DeleteFile baseDir+scoutSaveFile
end if
Set objectsFileToWrite = Nothing

Sub GetExposureData(expTime,imageCount)
	imageScale = 1.29
	skySeeing = 4
	call getObjectRate(object, objectRate)
	expTime = round((60*(imageScale/objectRate)*skySeeing),0)
	
	If (expTime >= 30) AND (expTime < 45) Then
		expTime = 30 
	ElseIf	(expTime >= 45) AND (expTime < 60) Then 
		expTime = 45 
	ElseIf expTime >= 60 Then 
		expTime = 60 
	End If
	
	imageCount = round((60*(60/expTime)),0)
	msgbox expTime & " " & imageCount
End Sub

Sub getObjectRate(object, objectRate)
	Dim objTheSkyChart
	Dim objTheSkyInfo
	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	status = objTheSkyChart.Find ("MPL "+object)
	Set objTheSkyInfo = CreateObject("TheSkyX.sky6ObjectInformation") 
	objTheSkyInfo.Index = 0 
	status = objTheSkyInfo.Property (77)
	objectRateRA = objTheSkyInfo.ObjInfoPropOut
	status = objTheSkyInfo.Property (78)
	objectRateDEC = objTheSkyInfo.ObjInfoPropOut
	objectRate = round(sqr((objRateRA*objRateRA)+(objectRateDEC*objectRateDEC))*60,2)
	Set objTheSkyChart = Nothing
	Set objTheSkyInfo = Nothing
End Sub
