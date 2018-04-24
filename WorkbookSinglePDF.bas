Sub WorkBookSinglePDFStart()

Dim wsA     As Worksheet
Dim wbA     As Workbook
Dim strTime As String
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim strPathFile As String
Dim myFile  As Variant
Dim WS_Count As Integer
Dim I       As Integer
Dim spltEndH As String
Dim strSpcPer As String

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
Dim sDocType As String
Dim sDirPath As String
Dim spltDayEnd As String
Dim spltYearNew As String

' Set WS_Count equal to the number of worksheets in the active workbook.
Set wbA = ActiveWorkbook
WS_Count = wbA.Worksheets.Count


 'Look for number in cell C1 or B1, split it by dash, display message if not there
    If Not IsEmpty(Range("C1")) Then
        sNumber = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
    ElseIf Not IsEmpty(Range("B1")) Then
        sNumber = LTrim(RTrim(Split(Range("B1").Value, "-")(0)))
     Else
        MsgBox ("No Document Number Found in WoorkBookSinglePDF, so cannnot run Macro...In the future you will be able to type your Number In.")
    End If
    
     'Look for Doctype in cell C1 or B1, split it by dash, display message if not there
    If Not IsEmpty(Range("C1")) Then
        sDocType = LTrim(RTrim(Split(Range("C1").Value, "-")(1)))
    ElseIf Not IsEmpty(Range("B1")) Then
        sDocType = LTrim(RTrim(Split(Range("B1").Value, "-")(1)))
     Else
        MsgBox ("No Document Number Found in WoorkBookSinglePDF, so cannnot run Macro...In the future you will be able to type your Number In.")
    End If

    sDate = Range("A1").Value
'    sNumber = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
'    sDocType = LTrim(RTrim(Split(Range("C1").Value, "-")(1)))
    
    sDateFinal = Replace(sDate, "/", ".")
    sNumberFinal = Replace(sNumber, ".", "")
    
  sDirPath = "file path"
    strPath = sDirPath & sDocType
    

    If Dir(sDirPath + sDocType, vbDirectory) = "" Then
        MkDir Path:=sDirPath + sDocType
       ' MsgBox "Created A New Folder:" + " " + sDocType + " " + "in the GL Reconciliation"
     End If
        
    If strPath = "" Then
        strPath = Application.DefaultFilePath
    End If
    
    strPath = strPath & "\"
    
    
    '----Split Date----
    spltMonth = Split(sDateFinal, ".")(0)
    spltDay = Split(sDateFinal, ".")(1)
    spltYear = Split(sDateFinal, ".")(2)
    spltSpace = Split(spltDay, " ")(0)
   
   
   '----Add Leading Zero-----
    If Len(spltMonth) < 2 Then
        spltMonthNew = "0" + spltMonth
        
    ElseIf Len(spltMonth) >= 2 Then
        spltMonthNew = spltMonth
    End If
    
    If Len(spltSpace) < 2 Then
        spltDayNew = "0" + spltDay
        
    ElseIf Len(spltSpace) >= 2 Then
        spltDayNew = spltDay
    End If
    
    
    '----Add Digits to Yr-----
    If Len(spltYear) < 4 Then
        spltYearNew = "20" + spltYear
        
    ElseIf Len(spltYear) >= 4 Then
        spltYearNew = spltYear
    End If
    
    
    If Len(sNumberFinal) > 10 Then
        sNumberFinal = "MULTIPLE"
    End If
    
    
    
    
    'add day and month back together
    'use date on worksheet
    'strName = spltMonthNew + "." + spltDayNew + "." + spltYearNew + " " + sNumberFinal + " " + sDocType
    
    'use end of month day in format YYYY.mm.dd + No DocType
    strName = spltYearNew + "." + spltMonthNew + "." + spltDayNew + " " + sNumberFinal + " " + sDocType
    
    
    
    'replace spaces and periods in sheet name
    'strNameNew = Replace(strName, " ", ".")
    strNameNew = strName

    'create default name for savng file
    strFile = strNameNew & ".pdf"
    myFile = strPath & strFile
    
     Application.PrintCommunication = False
            For Each wsA In wbA.Worksheets
                With wsA.PageSetup
                    .PrintArea = ""
                    .Orientation = xlLandscape
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End With
            Next wsA
    Application.PrintCommunication = True


    Debug.Print myFile

    'export to PDF if a folder was selected
        If myFile <> "False" Then
        wbA.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                fileName:=myFile, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
  
    End If

 ActiveWorkbook.FollowHyperlink Address:=strPath, NewWindow:=True

End Sub
