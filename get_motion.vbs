
Dim Elements
'Elements = "A10b8bg 22.9  0.15  K18CS   6.25119  306.65090  107.13873    7.88364  0.6610669  0.23391488   2.6087401                 26   1    1 days 0.39         NEOCPNomin           A10b8bg"
'Elements ="C00GHM2 22.9  0.15  K18CS  23.16628  114.68489  283.79415   31.09145  0.4065578  0.54602002   1.4825079                 38   1    1 days 0.33         NEOCPNomin"
Elements = "00001    3.34  0.12 K194R  77.37215   73.59764   80.30553   10.59407  0.0760091  0.21388522   2.7691652  0 MPO452155  6725 114 1801-2018 0.60 M-v 30h MPCLINUX   0000      (1) Ceres              20180430"
'Elements = "C00D322 25.9  0.15  K18CS 238.60491  309.29204  286.80578   17.22865  0.2152832  1.14110070   0.9069571                 25   1    1 days 0.62         NEOCPNomin"
'Elements = "D4340   -0.76  0.12 K194R  41.54710  115.01131  110.29733   17.10337  0.2537532  0.00391855  39.8461085  0 MPO459634  9482  94 1914-2018 0.30 M-v 38h MPCLINUX   000A (134340) Pluto              20180927"
Set Util = CreateObject("ASCOM.Utilities.Util")
Set AstroUtils = CreateObject("ASCOM.Astrometry.NOVAS.NOVAS31")
'Set Planet = CreateObject("ASCOM.Astrometry.Kepler.Ephemeris")

Call MinorPlanet(Elements, Name, RightAscension, Declination, RightAscensionRate, DeclinationRate)

Wscript.Echo " MP elements for " & Name & " --> RARate=" & _
               Round(RightAscensionRate,6) & " DecRate=" & _
                Round(DeclinationRate,6) & " " 
	Wscript.Echo " MP elements for " & Name & " --> RA=" & _
               RightAscension & " Dec=" & _
                Declination 
objectRate = Round(sqr((RightAscensionRate*RightAscensionRate) + (DeclinationRate*DeclinationRate))*60,6)
	Wscript.Echo " MP elements for " & Name & " --> Rate=" & objectRate & " Arcsec/Min"
	
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
		'Wscript.Echo "eqdist=" & round(eqdist,6)
		If eqdist > 0.0083334 Then
		     
            ' Moves "enough", calculate coordinate rates
            x = tvec2.RightAscension - tvec1.RightAscension
			Wscript.Echo "RAmoved=" & round(x,6) & " iterations=" & round(i)
			
            If x < -12.0 Then x = x + 12.0
            If x > 12.0 Then x = x - 12.0
			
            'RADot = (60 * x) / (dt * 86400.0)               ' RA coordinate rate sec/sec
			 RADot = x * j
				
			'Wscript.Echo "RA2=" & tvec2.RightAscension & " dt=" & round(dt,6) & " x=" & round(x,6) & " RADot=" & round(RADot,6)
			
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