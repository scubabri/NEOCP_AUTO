Set objShell = CreateObject("WScript.Shell")

Set Util = CreateObject("ASCOM.Utilities.Util")
Set AstroUtils = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")

' set your working directory here
baseDir = "E:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"  

' set your minimums here
minscore = 0
mindec = 0 
minvmag = 20	
minobs = 4
minseen = .8

strScriptFile = Wscript.ScriptFullName 													' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO\parse_neo.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strScriptFile)
strFolder = objFSO.GetParentFolderName(objFile) 										' D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO
 
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

objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(neocpLink) & " -N",0,True 	'download current neocp.txt from MPC 
 
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
	
	Wscript.Echo object & "     " & score & "  " + ra & "  " & dec & "    " & vmag & "      " & obs & "     " & seen
		
		objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(baseDir) & orbSaveFile,0,True ' run wget to get orbits from NEOCP
		
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
	   
	   objShell.Run Quotes(strFolder & "\wget.exe") & " " & Quotes(orbLinkBase) & object & "&orb=y -O" & " " & Quotes(baseDir) & orbSaveFile,0,True ' run wget to get orbits from NEOCP
		
		Set objRegEx = CreateObject("VBScript.RegExp")									' This section parses the orbits from the neocp object and gets the NEOCPNomin parameters
		objRegEx.Pattern = "NEOCPNomin"													' which is used in the MPCORB.dat 
        Set objFile = objFSO.OpenTextFile(baseDir+orbSaveFile, 1)
		Do Until objFile.AtEndOfStream
			strSearchString = objFile.ReadLine
			Set colMatches = objRegEx.Execute(strSearchString)
			
			If colMatches.Count > 0 Then
					mpcLine = strSearchString+"           "+object	'set MPCorb ephemeris for import to ACP
					Elements = strSearchString
				Exit Do
			End If
		Loop
	
		Name = object
		call GetExposureData(expTime,imageCount, objectRate)
		Call MinorPlanet(Elements, Name, RightAscension, Declination, RightAscensionRate, DeclinationRate)
		objectRateSCR = Round(sqr((RightAscensionRate*RightAscensionRate) + (DeclinationRate*DeclinationRate))*60,6)
		
		If imageCount > 0 Then											' to overcome issues when object has been moved to PCCP
		
			Dim RTML, REQ, TGT, PIC, COR, FSO, FIL, TR                    ' (for PrimalScript IntelliSense)

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
			Set TR = CreateObject("DC3.RTML23.TimeRange")
			
			'TR.Earliest = Cdate("13:45:00")
			
			'REQ.Schedule = TR
			Set REQ.Correction = CreateObject("DC3.RTML23.Correction")
			REQ.Correction.zero = False
			REQ.Correction.flat = False
			REQ.Correction.dark = False
		
			RTML.RequestsC.Add REQ
			REQ.ID = object                                         ' This becomes the Plan name for the Request
			REQ.Description = object + " Score: " + score + " RA: " + ra + " DEC: " + dec + " vMag: " + vmag + " #Obs: " + obs + " Last Seen: " + seen  + " Rate: " + CStr(objectRate) + " SCRRate: " + Cstr(Round(objectRateSCR,2)) + " arcsec/min"
			Set TGT = CreateObject("DC3.RTML23.Target")
			TGT.TargetType.OrbitalElements = mpcLine
			TGT.Description = mpcLine
			TGT.Name = object
			TGT.count = 2
			TGT.Interval= 2  ' HOURS!!!
			'TGT.Timefromprev = 0.5 ' hours
			'TGT.Tolfromprev = 0 
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
Set objTheSky = CreateObject("TheSkyX.sky6RASCOMTheSky") 
call objTheSky.Quit()
Set objTheSky = nothing	

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

Sub GetExposureData(expTime,imageCount, objectRate)
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
	
	imageCount = round((30*(60/expTime)),0)
End Sub

Sub getObjectRate(object, objectRate)
	Dim objTheSkyChart
	Dim objTheSkyInfo
	Set objTheSkyChart = CreateObject("TheSkyX.sky6StarChart") 
	status = objTheSkyChart.Find ("MPL "+object)
	Set objTheSkyInfo = CreateObject("TheSkyX.sky6ObjectInformation") 
	objTheSkyInfo.Index = 0 
	status = objTheSkyInfo.Property (77)
	objectRateRA = Cdbl(objTheSkyInfo.ObjInfoPropOut)
	status = objTheSkyInfo.Property (78)
	objectRateDEC = CDbl(objTheSkyInfo.ObjInfoPropOut)
	objectRate	= round((sqr((objectRateRA*objectRateRA)+(objectRateDEC*objectRateDEC)))*60,2)

	Set objTheSkyChart = Nothing
	Set objTheSkyInfo = Nothing
End Sub
	
Function MinorPlanet(Elements, Name, RightAscension, Declination, RightAscensionRate, DeclinationRate)
    
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
    
    Name = Trim(Left(Elements, 7))                          ' Object name (return)
    kt.Name = Name
    kt.Epoch = PackedToJulian(Trim(Mid(Elements, 21, 5))) + 1 ' Epoch of osculating elements
    cl = GetLocale()                                        ' Get locale (. vs , ****)
    SetLocale "en-us"                                       ' Make sure numbers convert properly
    kt.M = CDbl(Trim(Mid(Elements, 27, 9)))                 ' Mean anomaly
	kt.n = CDbl(Trim(Mid(Elements, 81, 11)))                ' Mean daily motion (deg/day)
	kt.a = CDbl(Trim(Mid(Elements, 93, 11)))                ' Semimajor axis (AU)
	kt.e = CDbl(Trim(Mid(Elements, 71, 9)))                 ' Orbital eccentricity
	kt.Peri = CDbl(Trim(Mid(Elements, 38, 9)))              ' Arg of perihelion (J2000, deg.)
	kt.Node = CDbl(Trim(Mid(Elements, 49, 9)))              ' Long. of asc. node (J2000, deg.)
    kt.Incl = CDbl(Trim(Mid(Elements, 60, 9)))              ' Inclination (J2000, deg.)
	SetLocale cl                                            ' Restore locale
	
    jd = Util.JulianDate                                     ' Get current jd
	
    pl.DeltaT = AstroUtils.DeltaT(jd)                                 ' Delta T for NOVAS and Kepler
    
	Call GetPositionAndVelocity(pl, Site,  jd - (pl.DeltaT / 86400), RightAscension, Declination, RightAscensionRate, DeclinationRate)	
	
    Set pl = Nothing                                            ' Releases both Ephemeris objs
    MinorPlanet = True                                          ' Success
    
End Function

' GetPosVel() - Compute position and velocity (coordinate rates) of solar system body
' -----------
'
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
    
    dt = (5.0 / 1440.0)                                         ' Start with 5 minute interval
   
	Set tvec1 =  pl.GetTopocentricPosition(TJD, st, False)                 ' Get current position
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

Function PackedToJulian(Packed) 
    Dim yr, mo, dy, PCODE, YCODE
    PCODE = "123456789ABCDEFGHIJKLMNOPQRSTUV"
	YCODE = "IJK"
	
    yr = (17 + InStr(YCODE, Left(Packed, 1))) * 100             ' Century
    yr = yr + CInt(Mid(Packed, 2, 2))                           ' Year in century   
    mo = InStr(PCODE, Mid(Packed, 4, 1))                        ' Month (1-12)
    dy = CDbl(InStr(PCODE, Mid(Packed, 5, 1)))                  ' Day (1-31)
    
	Call DateToJulian(yr, mo, dy, dtj)                   ' UTC Julian Date
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

Function ArcCos(X)
    ArcCos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function