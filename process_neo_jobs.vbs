Set objShell = CreateObject("WScript.Shell")

Set Util = CreateObject("ASCOM.Utilities.Util")
Set AstroUtils = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")

baseDir = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"  										' set your working directory here

Dim Elements, minscore, mindec, minvmag, minobs, minseen, imageScale, skySeeing, maxObjectMove, imageOverhead, binning, minHorizon, uncertainty, object, score, ra, dec, vmag, obs, seen

' set your minimums here
minscore = 0																			' what is the minumum score from the NEOCP, higher score, more desirable for MPC, used for Scheduler priority as well.
mindec = -10 																			' what is the minimum dec you can image at
minvmag = 19.5																			' what is the dimmest object you can see
minobs = 3																				' how many observations, fewer observations mean chance of being lost
minseen = 10																				' what is the oldest object from the NEOCP, older objects have a good chance of being lost.
imageScale = 1.29																		' your imageScale for determining exposure duration for moving objects
skySeeing = 4																			' your skyseeing in arcsec, used for figuring out max exposure duration for moving objects.
imageOverhead = 5 																		' how much time to download (and calibrate) added to exposure duration to calculate total number of exposures and repoint
maxObjectMove = 10 																		' this is the maximum we would like the object to move before repoint.
binning = 2 																			' binning
minHorizon = 30																			' minimum altitude that ACP/Scheduler will start imaging
maxuncertainty = 100																		' maximum uncertainty in arcmin from scout for attempt 
getMPCORB = True																		' do you want the full MPCORB.dat for reference, new NEOCP objects will be appended.
getCOMETS = True
getNEOCP = True
getESAPri = True

strScriptFile = Wscript.ScriptFullName 													
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
runDir = objFSO.GetParentFolderName(objFile) 
objShell.CurrentDirectory = baseDir        												' E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO

esaPlistLink = "http://neo.ssa.esa.int/PSDB-portlet/plist.txt"							' link to the ESA Priority List.
mpcLinkBase = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="			' base url to get NEOCPNomin orbit elements, shouldnt need to change this
neocpLink = "https://minorplanetcenter.net/iau/NEO/neocp.txt" 							' minorplanetcenter URL, shouldnt need to change this
scoutLink = "https://ssd-api.jpl.nasa.gov/scout.api?tdes="								' link to jpl scout for uncertainty determination
newtonLinkBase = "https://newton.spacedys.com/~neodys2/mpcobs/"

neocpTmpFile = baseDir+"\neocp.txt"														' where to put the downloaded neocp.txt, adjust as required.
objectsSaveFile = baseDir+"\object_run.txt"												' where to put output of selected NEOCP objects reference
scoutTmpFile = "\scout.json"															' temporary scout json file
plistTmpFile = basedir+"\plist.txt"																' temporary plist file

mpcTmpFile = "\find_o64\mpc_fmt.txt"													' this is the output from find_orb in mpc 1-line element format
obsTmpFile = "\find_o64\observations.txt"												' this is from observations from the NEOCP after parsing and filtering NEOCP.txt

mpcorbSaveFile = baseDir+"\MPCORB.dat"	
												' the raw MPCORB.dat from MPC that we'll append our elements to.
fullMpcorbSave = "C:\Program Files (x86)\Common Files\ASCOM\MPCORB\MPCORB.dat"			' this is a copy for ACP should we decide to manually do an object run. 
fullMpcorbLink = "https://minorplanetcenter.net/iau/MPCORB/MPCORB.DAT"	
fullMpcorbDat = "\MPCORB.dat"

fullCometSave = "C:\Program Files (x86)\Common Files\ASCOM\MPCCOMET\CometEls.txt"		' this is a copy for ACP should we decide to manually do an object run. 
cometsLink = "https://minorplanetcenter.net/iau/MPCORB/CometEls.txt"
fullCometDat = "\CometEls.txt"															' I cant remember why I have this one and I'm too tired to figure it out.	

Include "VbsJson"	
Dim json, neocpStr, jsonDecoded
Set json = New VbsJson

if objFSO.FileExists(mpcorbSaveFile) then												' remove the old MPCORB.dat 
		objFSO.DeleteFile mpcorbSaveFile
end if
	
if objFSO.FileExists(objectsSaveFile) then												' remove the old object_run.txt 
	objFSO.DeleteFile objectsSaveFile
end if

call downloadObjects()
call updateACPObjects()	

If getNEOCP = True Then
	call getNEOCPObjects()
End If 

If getESAPri = True Then
	call getESAObjects()
End If

if objFSO.FileExists(baseDir+mpcTmpFile) then												' clean up temporary files
	objFSO.DeleteFile basedir+mpcTmpFile
end if

if objFSO.FileExists(baseDir+"\NEOCP.rtml") then											' clean up temporary files
	objFSO.DeleteFile baseDir+"\NEOCP.rtml"
end if

if objFSO.FileExists(neocpTmpFile) then
	objFSO.DeleteFile neocpTmpFile
end if
if objFSO.FileExists(baseDir+scoutTmpFile) then
	objFSO.DeleteFile baseDir+scoutTmpFile
end if
Set objFSO = Nothing
Set objectsFileToWrite = Nothing
																													
Function downloadObjects()

	if getCOMETS = true Then
		Wscript.Echo "Downloading CometEls.txt...."
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(cometsLink) & " -O" & " " & Quotes(baseDir) & fullCometDat,0,True  	' get the full comets file for reference
		Wscript.Echo "Done"
	End If

	if getMPCORB = true Then
		Wscript.Echo "Downloading MPCORB.dat...."
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(fullMpcorbLink) & " -O" & " " & Quotes(baseDir) & fullMpcorbDat,0,True  	' get the full MPCORB.dat file for reference
		Wscript.Echo "Done"
	End If

	objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(neocpLink) & " -N",0,True 									'download current neocp.txt from MPC 
	objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(esaPlistLink) & " -N",0,True								'download ESA Priority List
	
End Function

Function getESAObjects()

	set plistTmpFileRead = objFSO.OpenTextFile(plistTmpFile, 1) 
	Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").openTextFile(objectsSaveFile,8,true)  		' create object_run.txt
	
	if objFSO.FileExists(mpcorbSaveFile) then												
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat output to append NEOCP elements
	else
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").CreateTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat didnt exist for some reason, lets create and empty one
	End If
	
	Do Until plistTmpFileRead.AtEndOfStream													' read the downloaded neocp.txt and parse for object parameters
		strLine = plistTmpFileRead.ReadLine													' its probably a good idea NOT to touch the positions as they are fixed position.
		esaPriority = Mid(strLine,1,1)
		object = replace(Mid(strLine, 5,8),chr(34), chr(32))															' temporary object designation
		dec = Mid(strLine, 27,5)															' declination 
		vmag = Mid(strLine, 37,4)															' vMag 
		uncertainty = Mid(strLine, 42,5)
		
		If IsNumeric(esaPriority) Then
			If (Csng(dec) >= mindec) AND (Csng(vmag) <= minvmag) AND (Csng(uncertainty) <= maxuncertainty) Then
				wscript.echo object & " " & esaPriority & " " & dec & " " & vmag & " " & uncertainty
				objectsFileToWrite.WriteLine object + " " + esaPriority + " " + dec & " " + vmag + " " + uncertainty
				objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(newtonLinkBase) & Replace(object," ","")  & ".rwo" & " -O" & " " & Quotes(baseDir) & obsTmpFile,0,True 
				Set objFile = objFSO.OpenTextFile(baseDir+obsTmpFile, 1)
				
				If (objFile.AtEndOfStream <> True) Then
					objShell.CurrentDirectory = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\find_o64"		' lets change our cwd to run find_orb, it likes it's home
					objShell.Run "find_o64.exe observations.txt",0,False							' open the mpc_fmt.txt that find_orb created and append it to the MPCORB.dat
					Wscript.Sleep 5000																' find_orb needs a little bit to run, we'll give it 5 seconds.
					objShell.CurrentDirectory = baseDir												' change directory back to where we are running script
					objShell.Run "taskkill /im find_o64.exe"										' kill find_o64 when we are done until we get fo.exe for batch.
					Set objFile = objFSO.OpenTextFile(baseDir+mpcTmpFile, 1)
					Do Until objFile.AtEndOfStream
						Elements = objFile.ReadLine
						MPCorbFileToWrite.WriteLine(Elements)										'write elemets to MPCORB.dat
						call buildObjectDB(object)
					Loop
					objFile.Close
				End If
			End If
		End If
	Loop
End Function

Function getNEOCPObjects()
	
	Set neocpTmpFileRead = objFSO.OpenTextFile(neocpTmpFile, 1) 														' change path for input file from wget 
	Set objectsFileToWrite = CreateObject("Scripting.FileSystemObject").CreateTextFile(objectsSaveFile,8,true)  		' create object_run.txt

	if objFSO.FileExists(mpcorbSaveFile) then												
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat output to append NEOCP elements
	else
		Set MPCorbFileToWrite = CreateObject("Scripting.FileSystemObject").CreateTextFile(mpcorbSaveFile,8,true)  		' MPCORB.dat didnt exist for some reason, lets create and empty one
	End If
	
	Do Until neocpTmpFileRead.AtEndOfStream													' read the downloaded neocp.txt and parse for object parameters
		strLine = neocpTmpFileRead.ReadLine													' its probably a good idea NOT to touch the positions as they are fixed position.
		object = Mid(strLine, 1,7)															' temporary object designation
		score = Mid(strLine, 9,3)															' neocp desirablility score from 0 to 100, 100 being most desirable.
		ra	  = Mid(strLine, 27,7)															' right ascension 
		dec = Mid(strLine, 35,6)															' declination 
		vmag = Mid(strLine, 44,4)															' vMag 
		obs = Mid(strLine, 79,4)															' how many observations has it had
		seen = Mid(strLine, 96,7)															' when was the object last seen
	
		objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(scoutLink) & object & " -O" & " " & Quotes(baseDir) & scoutTmpFile,0,True ' Get NEOCP from Scout
		scoutStr = objFSO.OpenTextFile(baseDir+scoutTmpFile).ReadAll
		Set jsonDecoded = json.Decode(scoutStr)
		uncertainty = jsonDecoded("unc")												' position uncertainty in arcmin
	
		if (CSng(score) >= minscore) AND (CSng(dec) >= mindec) AND (CSng(vmag) <= minvmag) AND (CSng(obs) >= minobs) AND (CSng(seen) <= minseen) AND ((CSng(uncertainty) <= maxuncertainty) AND (uncertainty <> "")) Then
	
			objShell.Run Quotes(runDir & "\wget.exe") & " " & Quotes(mpcLinkBase) & object & "&obs=y -O" & " " & Quotes(baseDir) & obsTmpFile,0,True ' run wget to get observations from NEOCP
			Set objFile = objFSO.OpenTextFile(baseDir+obsTmpFile, 1)
			objFile.ReadAll
			if objFile.Line > 3 Then
				objShell.CurrentDirectory = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\find_o64"		' lets change our cwd to run find_orb, it likes it's home
				objShell.Run "find_o64.exe observations.txt",0,False							' open the mpc_fmt.txt that find_orb created and append it to the MPCORB.dat
				Wscript.Sleep 5000																' find_orb needs a little bit to run, we'll give it 5 seconds.
				objShell.CurrentDirectory = baseDir												' change directory back to where we are running script
				objShell.Run "taskkill /im find_o64.exe"										' kill find_o64 when we are done until we get fo.exe for batch.
				Set objFile = objFSO.OpenTextFile(baseDir+mpcTmpFile, 1)
				Do Until objFile.AtEndOfStream
					Elements = objFile.ReadLine
					MPCorbFileToWrite.WriteLine(Elements)										'write elemets to MPCORB.dat
				Loop
				objFile.Close
		
				Wscript.Echo object & "     " & score & "  " + ra & "  " & dec & "    " & vmag & "      " & obs & "     " & seen & "   " & uncertainty
				objectsFileToWrite.WriteLine(object+ "     " + score + "  " + ra + "  " + dec + "    " + vmag + "      " + obs + "     " + seen + " " + uncertainty)	' write out the objects_run for reference
				Name = object
				
				call buildObjectDB(object)
			End If
		End If
		
	Loop
	
	neocpTmpFileRead.Close																			' close any open files
	objectsFileToWrite.Close
	MPCorbFileToWrite.Close
End Function

Sub GetExposureData(expTime,imageCount, objectRate)
	
	call getMinorPlanetMotion(Elements, Name, RightAscension, Declination, RightAscensionRate, DeclinationRate)
	objectRate = Round(sqr((RightAscensionRate*RightAscensionRate) + (DeclinationRate*DeclinationRate))*60,2)
	expTime = round((60*(imageScale/objectRate)*skySeeing),0)
	
	If (expTime >= 30) AND (expTime < 45) Then
		expTime = 45 
	ElseIf	(expTime >= 45) AND (expTime < 60) Then 
		expTime = 60
	ElseIf expTime >= 60 AND (expTime < 75) Then 
		expTime = 75 
	ElseIf expTime >= 75 AND (expTime < 90) Then 
		expTime = 90
	ElseIf expTime >= 90 AND (expTime < 105) Then 
		expTime = 105
	ElseIf expTime >= 105 AND (expTime < 120) Then 
		expTime = 120
	ElseIf expTime >= 120  Then
		expTime = 120
	End If
	Minutes = 30 												' set to how long you want to capture images
	imageCount = round((Minutes*(60/expTime)),0)
End Sub

Function buildObjectDB(object)

	Name = object
	call GetExposureData(expTime,imageCount, objectRate)
	
	If imageCount > 0 Then															' to overcome issues when object has been moved to PCCP
	
		Dim RTML, REQ, TGT, PIC, COR, FSO, FIL, TR                 
		Set RTML = CreateObject("DC3.RTML23.RTML")
		Set RTML.Contact = CreateObject("DC3.RTML23.Contact")
		RTML.Contact.User = "neocp"
		RTML.Contact.Email = "brians@fl240.com"
		
		Set REQ =  CreateObject("DC3.RTML23.Request")
		REQ.UserName = "neocp"                                  
		REQ.Project = "NEOCP"                                    					' Proj for above user will be created if needed

		Set REQ.Schedule = CreateObject("DC3.RTML23.Schedule")
		REQ.Schedule.Horizon = minHorizon
		REQ.Schedule.Priority = score
		Set REQ.Schedule.Moon = CreateObject("DC3.RTML23.Moon")
		Set REQ.Schedule.Moon.Lorentzian = CreateObject("DC3.RTML23.Lorentzian")
		REQ.Schedule.Moon.Lorentzian.Distance  = 15
		REQ.Schedule.Moon.Lorentzian.Width = 6
			
		Set REQ.Correction = CreateObject("DC3.RTML23.Correction")
		REQ.Correction.zero = False
		REQ.Correction.flat = False
		REQ.Correction.dark = False
		
		RTML.RequestsC.Add REQ
		REQ.ID = object                                         					' This becomes the Plan name for the Request
		REQ.Description = object + " Score: " + score + " RA: " + ra + " DEC: " + dec + " vMag: " + vmag + " #Obs: " + obs + " Last Seen: " + seen  + " Rate: " + CStr(objectRate) + " arcsec/min" + " Unc (arcmin): " + uncertainty
		Set TGT = CreateObject("DC3.RTML23.Target")
		TGT.TargetType.OrbitalElements = Elements
		TGT.Description = Elements
		TGT.Name = object
			
		imageTotalTime = (((expTime + imageOverhead) * imageCount)/60)				' in minutes including overhead for download, etc
		totalMove = ((objectRate * imageTotalTime)/60)    							' in arcmin		
			
		baseTargetCount = round(maxObjectMove / (imageTotalTime / totalMove),0) 
		if baseTargetCount < 1 Then
			TGT.count = 1
		Else 
			TGT.count = baseTargetCount + 10000
		End If
			
		TGT.Interval = 0
		REQ.TargetsC.Add TGT
		
		Set PIC = CreateObject("DC3.RTML23.Picture")
		PIC.Name = object+" Luminance"                               	
		PIC.ExposureSpec.ExposureTime = expTime
		PIC.Binning = binning
		PIC.Filter = "Luminance"
		PIC.Description = "#nopreview"
			
		if baseTargetCount < 1 Then 
			PIC.Count = imageCount
		Else 
			PIC.Count = round(imageCount / baseTargetCount,0)
		End If
			
		TGT.PicturesC.Add PIC
		
		XML = RTML.XML(True)
		
		Set FSO = CreateObject("Scripting.FileSystemObject")
		Set FIL = FSO.CreateTextFile(baseDir+"\NEOCP.rtml", True)    		' **CHANGE FOR YOUR SYSTEM**
		FIL.Write XML                                           			' Has embedded line endings
		FIL.Close
		
		Dim I, DB, R
		Set DB = CreateObject("DC3.Scheduler.Database")
		Call DB.Connect()
		Set I = CreateObject("DC3.RTML23.Importer")
		Set I.DB = DB
			
		I.Import baseDir+"\NEOCP.rtml"
		Set R = I.Projects.Item(0)
			
		R.Disabled = false
		R.Update()
		Set NewPlan = I.Plans.Item(0)
		NewPlan.Resubmit()
            		
		Call DB.Disconnect()
		Set REQ =  nothing
		Set RTML = nothing
		Set TGT = nothing
		set PIC = Nothing
	End If
End Function

Function getMinorPlanetMotion(Elements, Name, RightAscension, Declination, RightAscensionRate, DeclinationRate)
    
	Dim kt, ke, pl, jd, mp, key, cl, Site
    set Site = CreateObject("NOVAS.Site")
    Set pl = CreateObject("NOVAS.Planet")
    Set kt = CreateObject("Kepler.Ephemeris")
    Set ke = CreateObject("Kepler.Ephemeris")
	
	Site.Height = 1540.2
	Site.Longitude = -111.760981
	Site.Latitude = 40.450216
	
    pl.Ephemeris = kt                                           ' Plug in target ephemeris gen
    pl.EarthEphemeris = ke                                      ' Plug in Earth ephemeris gen
    pl.Type = 1                                                 ' NOVAS: Minor Planet (Passed to Kepler)
    pl.Number = 1                                               ' Must pass valid number to Kepler, but is ignored
    
    Name = Trim(Left(Elements, 7))                          	' Object name (return)
    kt.Name = Name
    kt.Epoch = PackedToJulian(Trim(Mid(Elements, 21, 5))) + 1 	' Epoch of osculating elements
    cl = GetLocale()                                        	' Get locale (. vs , ****)
    SetLocale "en-us"                                       	' Make sure numbers convert properly
    kt.M = CDbl(Trim(Mid(Elements, 27, 9)))                 	' Mean anomaly
	kt.n = CDbl(Trim(Mid(Elements, 81, 11)))                	' Mean daily motion (deg/day)
	kt.a = CDbl(Trim(Mid(Elements, 93, 11)))                	' Semimajor axis (AU)
	kt.e = CDbl(Trim(Mid(Elements, 71, 9)))                 	' Orbital eccentricity
	kt.Peri = CDbl(Trim(Mid(Elements, 38, 9)))              	' Arg of perihelion (J2000, deg.)
	kt.Node = CDbl(Trim(Mid(Elements, 49, 9)))              	' Long. of asc. node (J2000, deg.)
    kt.Incl = CDbl(Trim(Mid(Elements, 60, 9)))             	 	' Inclination (J2000, deg.)
	SetLocale cl                                            	' Restore locale
	
    jd = Util.JulianDate                                     	' Get current jd
	
    pl.DeltaT = AstroUtils.DeltaT(jd)                            ' Delta T for NOVAS and Kepler
    
	Call GetPositionAndVelocity(pl, Site,  jd - (pl.DeltaT / 86400), RightAscension, Declination, RightAscensionRate, DeclinationRate)	
	
    Set pl = Nothing                                            ' Releases both Ephemeris objs
    MinorPlanet = True                                          ' Success
    
End Function

' GetPosVel() - Compute position and velocity (coordinate rates) of solar system body
' -----------
' This uses a cheapo linear extrapolation. Eventually, NOVAS needs to be updated
' to provide a VelocityVector for a Planet. The delta t is increased to provide at
' least 20 arcseconds movement. This hopefully avoids roundoff errors on slow moving 
' objects. It projects forward in time rather than bracketing the given time, since 
' the data will be taken in the future.
'
' pl      = [in]  NOVAS.Planet object
' TJD     = [in]  Terrestrial Julian Date for position
' RA      = [out] Right Ascension, hours
' Dec     = [out] Declination, degrees
' RADot   = [out] RightAscension Rate, seconds per Second
' DecDot  = [out] Declination Rate, arcseconds per second
'----------------------------------------------------------------------------------------
Function GetPositionAndVelocity(pl, st, TJD, RA, Dec, RADot, DecDot)
    Dim dt, tvec1, tvec2, x, y, i
    
    dt = (5.0 / 1440.0)                                         	' Start with 5 minute interval
   
	Set tvec1 =  pl.GetTopocentricPosition(TJD, st, False)      	' Get current position
    RA = tvec1.RightAscension
    Dec = tvec1.Declination
	
    ' Keep doubling the interval until we get 30 arcsec total movement.
    ' If we don't get there in 13 or fewer steps (28 days), something is 
    ' wrong. We should get that on TNOs.
   
    j = 180 
    For i = 1 To 13  	' Goes out to 28 day interval
        Set tvec2 = pl.GetTopocentricPosition(TJD + dt, st, False)
        if i > 1 Then j = j / 2

		Call EquDist(tvec1.RightAscension, tvec1.Declination, tvec2.RightAscension, tvec2.Declination, eqdist)
		
		If eqdist > 0.0083334 Then
            ' Moves "enough", calculate coordinate rates
            x = tvec2.RightAscension - tvec1.RightAscension
			
            If x < -12.0 Then x = x + 12.0
            If x > 12.0 Then x = x - 12.0
			
            'RADot = (60 * x) / (dt * 86400.0)               ' RA coordinate rate sec/sec
			 RADot = x * j
				
            If Abs(x) > 6.0 Then                                ' Moved across pole
                y = tvec2.Declination + tvec1.Declination       ' Total dec movement is sum
                If Dec >= 0.0 Then                              ' Moved across north pole
                    y = 180.0 - y
                Else                                            ' Moved across south pole
                    y = -180.0 - y
                End If
            Else                                                ' Same side of pole
                y = tvec2.Declination - tvec1.Declination
            End If
            DecDot = (3600 * y) / (dt * 86400.0)                ' Dec rate arcsec/sec
			
            Set tvec1 = Nothing
            Set tvec2 = Nothing
			
            Exit Function                                       ' == SUCCESS, EXIT FUNCTION ==
        End If
        dt = 2.0 * dt
    Next
	
    Set tvec1 = Nothing
    Set tvec2 = Nothing
End Function

Function PackedToJulian(Packed) 								' https://minorplanetcenter.net/iau/info/PackedDates.html
    Dim yr, mo, dy, PCODE, YCODE
    PCODE = "123456789ABCDEFGHIJKLMNOPQRSTUV"
	YCODE = "IJK"
	
    yr = (17 + InStr(YCODE, Left(Packed, 1))) * 100             ' Century
    yr = yr + CInt(Mid(Packed, 2, 2))                           ' Year in century   
    mo = InStr(PCODE, Mid(Packed, 4, 1))                        ' Month (1-12)
    dy = CDbl(InStr(PCODE, Mid(Packed, 5, 1)))                  ' Day (1-31)
    
	Call DateToJulian(yr, mo, dy, dtj)                  		' UTC Julian Date
	PackedToJulian = dtj 		
End Function

' DateToJulian() - Convert Gregorian calendar date to Julian
'Good for Gregorian dates after 28-Feb-1900. The Util.Date_Julian() method in ACP is
'a pain to use because the date is Local time. 

function DateToJulian(yr, mo, dy, dtj)
    dtj = ((367 * yr) - Round((7 * (yr + Round((mo + 9) / 12))) / 4) + Round((275 * mo) / 9) + dy + 1721013.5)
End Function

' EquDist() - Return equatorial distance between objects, degrees
'---------
'
'----------------------------------------------------------------------------------------
function EquDist(a1, d1, a2, d2, eqdist)  
	a1b = (a1 * 15.0)
	a2b = (a2 * 15.0)

	Call SphDist((a1 * 15.0), d1, (a2 * 15.0), d2, spdist)
		eqdist = spdist	
End Function

' SphDist() - Return distance between objects, degrees 
Dim a1b, a2b

Function SphDist(a1b, d1, a2b, d2, spdist)	
	pi = 4 * atn(1.0)
	DEGRAD = pi / 180.0
	RADDEG = 180.0 / pi
	
    a1r = DEGRAD * a1b
    a2r = DEGRAD * a2b
    d1r = DEGRAD * d1
    d2r = DEGRAD * d2
      
    ca1 = cos(a1r)
    ca2 = cos(a2r)
    sa1 = sin(a1r)
    sa2 = sin(a2r)

    cd1 = cos(d1r)
    cd2 = cos(d2r)
    sd1 = sin(d1r)
    sd2 = sin(d2r)

    x1 = cd1*ca1
    x2 = cd2*ca2
    y1 = cd1*sa1
    y2 = cd2*sa2
    z1 = sd1
    z2 = sd2

    R = (x1 * x2) + (y1 * y2) + (z1 * z2)

    if(R > 1.0) Then 
		R = 1
	End If
	
    if(R < -1.0) Then 
		R = -1.0
	End If

    spdist = (RADDEG * arccos(R)) 
End Function

Function updateACPObjects()
	if getMPCORB = True Then
		set mpccopy=CreateObject("Scripting.FileSystemObject")
		mpccopy.CopyFile baseDir+fullMpcorbDat, fullMpcorbSave, True

		objShell.CurrentDirectory = "C:\Program Files (x86)\Common Files\ASCOM\MPCORB"
		objShell.Run "MakeDB.wsf",0,True

		set mpccopy = nothing
		objShell.CurrentDirectory = baseDir
	End If

	if getCOMETS = True Then
		set cmtccopy=CreateObject("Scripting.FileSystemObject")
		cmtccopy.CopyFile baseDir+fullCometDat, fullCometSave, True

		objShell.CurrentDirectory = "C:\Program Files (x86)\Common Files\ASCOM\MPCCOMET"
		objShell.Run "MakeCometDB.wsf",0,True

		set mpccopy = nothing
		objShell.CurrentDirectory = baseDir
	End If
End Function

Function ArcCos(X)
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Function Include(file)
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
End Function

Function Quotes(strQuotes)																' Add Quotes to string
	Quotes = chr(34) & strQuotes & chr(34)												' http://stackoverflow.com/questions/2942554/vbscript-adding-quotes-to-a-string
End Function
