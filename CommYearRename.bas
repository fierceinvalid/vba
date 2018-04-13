Sub CommYearOnly()

Dim strName As String
Dim strPath As String
Dim strFile As String
Dim myFile  As Variant

Dim spltDay As String
Dim spltDayNew As String
Dim spltMonth As String
Dim spltMonthNew As String
Dim spltSpace As String
Dim strNameNew As String
Dim sDate As String
Dim sDateFinal As String
Dim sNumber As String
Dim sNumberFinal As String
Dim spltYear As String
Dim spltYearNew As String
Dim g As String

Dim Path As String
Dim Filename As String
Dim xString As String
Dim xString2 As String
Dim xString3 As String
Dim spltFilename As String
Dim spltNew As String
Dim OldFile As String
Dim NewFile As String

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "\d{4}(\s|\w|-)"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True

    Dim colFiles As Collection
    Set colFiles = New Collection

    RecursiveFileSearch "put your file path here", objRegExp, colFiles, objFSO

    For Each f In colFiles
    'Debug.Print (f)
        'Insert code here to do something with the matched files
        
        Dim regex As Object, str As String
        Set regex = CreateObject("VBScript.RegExp")
 
            With regex
                .Pattern = "^\d{4}\s"
                .Global = True
            End With
     

        xString = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        Set matches = regex.Execute(xString)
        For Each Match In matches
        
        spltSpace = Split(xString, " ", 2)(0)
        spltFilename = Split(xString, " ", 2)(1)
        
        strYearDM = spltSpace + ".12." + "31"
        
        'Debug.Print Match.Value
        'Debug.Print strYearDM
        
           strName = strYearDM + " " + spltFilename
        g = Replace(f, xString, strName)
        strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
        Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
        

        
        Dim regex2 As Object, str2 As String
        Set regex2 = CreateObject("VBScript.RegExp")
 
            With regex2
                .Pattern = "^\d{4}\w"
                .Global = True
            End With
            
                

        xString2 = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        
        Set matches = regex2.Execute(xString2)
        For Each Match In matches
        
        
        
        
        spltSpace = Left(xString2, 4)
        spltFilename = Right(xString2, Len(xString2) - Len(spltSpace))
        'spltFilename = Split(xString2, "8", 2)(1)
        
        
         
        
        strYearDM = spltSpace + ".12." + "31"
        
        'h = Replace(f, strYearDM, xString2)
        
        'Debug.Print Match.Value
        'Debug.Print strYearDM
        
           strName = strYearDM + " " + spltFilename
        g = Replace(f, xString2, strName)
        strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
        Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
        
        
        
        
        
        
            Dim regex3 As Object, str3 As String
        Set regex3 = CreateObject("VBScript.RegExp")
 
            With regex3
                .Pattern = "^\d{4}-"
                .Global = True
            End With
            
                

        xString3 = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        
        Set matches = regex3.Execute(xString3)
        For Each Match In matches
        
        
        
        
        spltSpace = Left(xString3, 4)
        spltFilename = Right(xString3, Len(xString3) - Len(spltSpace))
        'spltFilename = Split(xString2, "8", 2)(1)
        
        
         
        
        strYearDM = spltSpace + ".12." + "31"
        
        'h = Replace(f, strYearDM, xString2)
        
        'Debug.Print Match.Value
        'Debug.Print strYearDM
        
           strName = strYearDM + " " + spltFilename
        g = Replace(f, xString3, strName)
        strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
        Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
              

        
    Next

    'Garbage Collection
    Set objFSO = Nothing
    Set objRegExp = Nothing

End Sub




Sub RecursiveFileSearch(ByVal targetFolder As String, ByRef objRegExp As Object, _
                ByRef matchedFiles As Collection, ByRef objFSO As Object)

    Dim objFolder As Object
    Dim objFile As Object
    Dim objSubFolders As Object

    'Get the folder object associated with the target directory
    Set objFolder = objFSO.GetFolder(targetFolder)

    'Loop through the files current folder
    For Each objFile In objFolder.Files
        If objRegExp.test(objFile) Then
            matchedFiles.Add (objFile)
        End If
    Next

    'Loop through the each of the sub folders recursively
    Set objSubFolders = objFolder.SubFolders
    For Each objSubfolder In objSubFolders
        RecursiveFileSearch objSubfolder, objRegExp, matchedFiles, objFSO
    Next

    'Garbage Collection
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objSubFolders = Nothing

End Sub
