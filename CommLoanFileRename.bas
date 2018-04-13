Sub CommLoanFileRenameStart()

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
Dim spltFilename As String
Dim spltNew As String
Dim OldFile As String
Dim NewFile As String

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "(\.)\d"
    objRegExp.IgnoreCase = True

    Dim colFiles As Collection
    Set colFiles = New Collection

    RecursiveFileSearch "put your file path here", objRegExp, colFiles, objFSO

    For Each f In colFiles
        'Insert code here to do something with the matched files
        
        xString = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        sDateFinal = Replace(xString, "-", ".")
        spltYear = Split(sDateFinal, ".")(0)
        spltMonth = Split(sDateFinal, ".")(1)
        spltDay = Split(sDateFinal, ".")(2)
        spltSpace = Split(spltDay, " ", 2)(0)
        spltFilename = Split(spltDay, " ", 2)(1)
   
        'compare character length of month to add leading zero
        If Len(spltMonth) < 2 Then
            spltMonthNew = "0" + spltMonth
        ElseIf Len(spltMonth) >= 2 Then
            spltMonthNew = spltMonth
         End If
    
        'compare character length of day to add leading zero
        If Len(spltSpace) < 2 Then
            spltDayNew = "0" + spltSpace
        ElseIf Len(spltSpace) >= 2 Then
            spltDayNew = spltSpace
        End If
    
        'compare character length of year to add "20"
        If Len(spltYear) < 4 Then
            spltYearNew = "20" + spltYear
        ElseIf Len(spltYear) >= 4 Then
            spltYearNew = spltYear
        End If
    
        strName = spltYearNew + "." + spltMonthNew + "." + spltDayNew + " " + spltFilename
        g = Replace(f, xString, strName)
        strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
        Name OldFile As NewFile
        
        Debug.Print g
        
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
