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
Dim i       As Integer
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
Dim LastRow As Long

' Set WS_Count equal to the number of worksheets in the active workbook.
Set wbA = ActiveWorkbook
WS_Count = wbA.Worksheets.Count



    
     'Look for Doctype in cell C1 or B1, split it by dash, display message if not there
    If Not IsEmpty(Range("C1")) Then
        sDocType = LTrim(RTrim(Split(Range("C1").Value, "-")(1)))
    ElseIf Not IsEmpty(Range("B1")) Then
        sDocType = LTrim(RTrim(Split(Range("B1").Value, "-")(1)))
     Else
        MsgBox ("No Document Number Found in WoorkBookSinglePDF, so cannnot run Macro...In the future you will be able to type your Number In.")
    End If
    
    'Look for number in cell C1 or B1, split it by dash, display message if not there
    If Not IsEmpty(Range("C1")) Then
        sNumber = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
    ElseIf Not IsEmpty(Range("B1")) Then
        sNumber = LTrim(RTrim(Split(Range("B1").Value, "-")(0)))
     Else
        MsgBox ("No Document Number Found in WoorkBookSinglePDF, so cannnot run Macro...In the future you will be able to type your Number In.")
    End If


    'Look for date in cell A1 or B1
    If Not IsEmpty(Range("A1")) Then
       sDate = Range("A1").Value
    ElseIf Not IsEmpty(Range("B1")) Then
       sDate = Range("B1").Value
    Else:
        sDate = InputBox _
        ("No Date was found in A1 or B1, Please Enter Date in format mm/dd/yyyy" _
        , "Date" _
        , Format(Now() _
        , "mm/dd/yyyy"))
        If IsDate(sDate) Then
    sDate = Format(CDate(sDate), "mm/dd/yyyy")
'    MsgBox sDate
  Else
    MsgBox "Wrong date format, Maco has ended GOODBYE"
     Exit Sub
  End If
    
        If Len(sDate) = "0" Or sDate = "dd/mm/yyyy" Then
            MsgBox "Date Not Chosen, Maco has ended GOODBYE"
        Exit Sub
    End If
End If



'    sDate = Range("A1").Value
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
    
'    ActiveSheet.UsedRange
'    LastRow = Cells.SpecialCells(xlLastCell).Row
    
'     Application.PrintCommunication = False
            For Each wsA In wbA.Worksheets
            
             Dim lR As Long, lC As Long
    
    With wsA
        lR = .Cells.Find("*", after:=Cells(Rows.Count, 1), _
            searchorder:=xlRows, searchdirection:=xlPrevious).Row
        lC = .Cells.Find("*", after:=Cells(1, Columns.Count), _
            searchorder:=xlColumns, searchdirection:=xlPrevious).Column
        .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(lR, lC)).Address
        .PageSetup.Orientation = xlLandscape
        'to immediately go to printout use: (else comment out)
'        .PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
            
        'to see preview first use: (else comment out)
        .PrintPreview
    End With
            
            
'              LastRow = Range("A:N").SpecialCells(xlCellTypeLastCell).Row
                LastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
                LastRow = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Row - 1
'              wsA.Activate
'              ActiveSheet.UsedRange
'              ActiveSheet.ResetAllPageBreaks

''              ActiveSheet.Cells.ClearFormats
'                With wsA.PageSetup
' '                   .PrintArea = "$A$2:$R$" & LastRow
''                    .PrintArea = LastRow
'                    .Orientation = xlLandscape
'                    .Zoom = False
'                    .FitToPagesWide = 1
'                    .FitToPagesTall = False
'                End With
            Next wsA
'    Application.PrintCommunication = True


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
