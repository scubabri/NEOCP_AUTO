Imports System.Net
Imports System.IO
Imports System.Text.RegularExpressions

Module GlobalVariables

    'set your  minimums here
    Public minscore As Integer = 100
    Public mindec As Integer = 0
    Public minvmag As Integer = 20
    Public minobs As Integer = 3
    Public minseen As Integer = 10

    Public baseDir As String = "D:\Dropbox\ASTRO\SCRIPTS\NEOCP_AUTO"

    Public neocpLink As String = "https://minorplanetcenter.net/iau/NEO/neocp.txt"                           ' minorplanetcenter URL, shouldnt need to change this
    Public neocpFile As String = baseDir + "\neocp.txt"                                                        ' where to put the downloaded neocp.txt, adjust as required.

    Public orbLinkBase As String = "https://cgi.minorplanetcenter.net/cgi-bin/showobsorbs.cgi?Obj="          ' base url to get NEOCPNomin orbit elements, shouldnt need to change this
    Public orbSaveFile As String = baseDir + "\orbits.txt"

    Public mpcorbSaveFile As String = baseDir + "\MPCORB.dat"                                                  ' the final (almost) MPCORB.dat 
    Public objectsSaveFile As String = baseDir + "\output.txt"                                                 ' where to put output of selected NEOCP objects for further parsing

End Module

Module Module1
    Sub Main()
        If File.Exists(neocpFile) Then
            File.Delete(neocpFile)
        End If

        If File.Exists(mpcorbSaveFile) Then
            File.Delete(mpcorbSaveFile)
        End If

        If File.Exists(objectsSaveFile) Then
            File.Delete(objectsSaveFile)
        End If

        getNEOCP()
        parseNeoCP()

        If File.Exists(orbSaveFile) Then
            File.Delete(orbSaveFile)
        End If

        If File.Exists(neocpFile) Then
            File.Delete(neocpFile)
        End If

    End Sub

    Sub getNEOCP()

        Try
            Dim fileReader As New WebClient()

            If Not (File.Exists(neocpFile)) Then
                fileReader.DownloadFile(neocpLink, neocpFile)
            End If

        Catch ex As HttpListenerException
            Console.WriteLine("Error accessing " + neocpLink + " - " + ex.Message)
        Catch ex As Exception
            Console.WriteLine("Error accessing " + neocpLink + " - " + ex.Message)
        End Try

    End Sub

    Sub getNeoOrbit(neoObject)

        Try
            Dim fileReader As New WebClient()

            fileReader.DownloadFile(orbLinkBase + neoObject & "&orb=y", orbSaveFile)
            Console.WriteLine("wrote " & orbSaveFile)

        Catch ex As HttpListenerException
            Console.WriteLine("Error accessing " + orbLinkBase + neoObject & "&orb=y" + " - " + ex.Message)
        Catch ex As Exception
            Console.WriteLine("Error accessing " + orbLinkBase + neoObject & "&orb=y" + " - " + ex.Message)
        End Try

    End Sub

    Sub parseNeoCP()

        Dim neocpFileRead As System.IO.StreamReader = New System.IO.StreamReader(neocpFile)
        Dim MPCorbFileToWrite As System.IO.StreamWriter = New System.IO.StreamWriter(mpcorbSaveFile, True)
        Dim objectsFileToWrite As System.IO.StreamWriter = New System.IO.StreamWriter(objectsSaveFile, True)


        Do Until neocpFileRead.EndOfStream
            Dim strLine = neocpFileRead.ReadLine                                                                   ' its probably a good idea NOT to touch the positions as they are fixed position.
            Dim neoObject As String = Mid(strLine, 1, 7)                                                           ' temporary object designation
            Dim neoScore As Integer = Mid(strLine, 9, 3)                                                            ' neocp desirablility score from 0 to 100, 100 being most desirable.
            Dim neoRa As Decimal = Mid(strLine, 27, 7)
            Dim neoDec As Decimal = Mid(strLine, 35, 7)                                                            ' declination 
            Dim neoVmag As Decimal = Mid(strLine, 44, 4)                                                           ' if you dont know what this is, change hobbies
            Dim neoObs As Integer = Mid(strLine, 79, 4)                                                            ' how many observations has it had
            Dim neoLastSeen As Decimal = Mid(strLine, 96, 7)                                                        ' when was the object last seen

            If (CSng(neoScore) >= minscore) And (CSng(neoDec) >= mindec) And (CSng(neoVmag) <= minvmag) And (CSng(neoObs) >= minobs) And (CSng(neoLastSeen) <= minseen) Then

                getNeoOrbit(neoObject)

                Dim neoOrbitFileRead As System.IO.StreamReader = New System.IO.StreamReader(orbSaveFile)

                Dim pattern As String = "NEOCPNomin"                                         ' This section parses the orbits from the neocp object and gets the NEOCPNomin parameters
                Dim objRegEx As Regex = New Regex(pattern, RegexOptions.None)                                     ' which is used in the MPCORB.dat

                Do Until neoOrbitFileRead.EndOfStream
                    Dim strSearchString = neoOrbitFileRead.ReadLine
                    Dim colMatches = objRegEx.Match(strSearchString)
                    If colMatches.Success Then
                        Console.WriteLine(strSearchString)                                      'echo selected MPCORB element for testing only
                        MPCorbFileToWrite.WriteLine(strSearchString + "           " + neoObject)    'write elemets to MPCORB.dat
                    End If
                Loop

                neoOrbitFileRead.Close()

                objectsFileToWrite.WriteLine(neoObject & "     " & neoScore & "  " & neoRa & "  " & neoDec & "    " & neoVmag & "      " & neoObs & "       ")
            End If
        Loop

        objectsFileToWrite.Close()
        MPCorbFileToWrite.Close()
        neocpFileRead.Close()
    End Sub
End Module