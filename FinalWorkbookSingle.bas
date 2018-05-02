Option Explicit
Sub WorkBookSinglePDFStart()

Dim wsA     As Worksheet
Dim wbA     As Workbook
Dim strName As String
Dim strPath As String
Dim strFile As String
Dim myFile  As Variant
Dim WS_Count As Integer
Dim spltDay As String
Dim spltDayNew As String
Dim spltMonth As String
Dim spltMonthNew As String
Dim spltSpace As String
Dim sDate As String
Dim sDateFinal As String
Dim sNumber As String
Dim sNumberFinal As String
Dim spltYear As String
Dim sDocType As String
Dim sDirPath As String
Dim spltDayEnd As String
Dim LastRow As Long
Dim lR As Long
Dim lC As Long

    '----Set WS_Count equal to the number of worksheets in the active workbook----
    Set wbA = ActiveWorkbook
    WS_Count = wbA.Worksheets.Count
    
    
    '----Look for Doctype & Number, split it by dash, display message if not there----
    If Not IsEmpty(Range("C1")) Then
        sDocType = LTrim(RTrim(Split(Range("C1").Value, "-")(1)))
        sNumber = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
    ElseIf Not IsEmpty(Range("B1")) Then
        sDocType = LTrim(RTrim(Split(Range("B1").Value, "-")(1)))
        sNumber = LTrim(RTrim(Split(Range("B1").Value, "-")(0)))
     Else
        MsgBox ("No Document Number Found in WoorkBookSinglePDF, so cannnot run Macro...In the future you will be able to type your Number In.")
    End If
    

    '----Look for date in cell A1 or B1----
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
        Else
            MsgBox "Wrong Date Format, Maco has ended"
            Exit Sub
        End If
        
        If Len(sDate) = "0" Or sDate = "dd/mm/yyyy" Then
            MsgBox "Wrong Date Format, Maco has ended"
            Exit Sub
        End If
    End If


    sDateFinal = Replace(sDate, "/", ".")
    sNumberFinal = Replace(sNumber, ".", "")
    sDirPath = "C:\Users\dbalk\Desktop\Work\Eforms\Accounting\GL Reconciliation\"
    strPath = sDirPath & sDocType
    
    '----Make Correct Folder in Directroy----
    If Dir(sDirPath + sDocType, vbDirectory) = "" Then
        MkDir Path:=sDirPath + sDocType
     End If
    '----If no file Path is found save to GL Rec Main Folder
    If strPath = "" Then
'        strPath = Application.DefaultFilePath
         strPath = sDirPath
         MsgBox "No Folder could be made for" + sDocType + "so file will be saved to the GL Reconcilation Folder"
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
        spltYear = "20" + spltYear
    End If
    
    '----Check length of Doctype Number-----
    If Len(sNumberFinal) > 10 Then
        sNumberFinal = "MULTIPLE"
    End If
    
    '----Creat Name for saving file-----
    strName = spltYear + "." + spltMonthNew + "." + spltDayNew + " " + sNumberFinal + " " + sDocType
    strFile = strName & ".pdf"
    myFile = strPath & strFile
    
    '---Go through each worksheet and correct format----
    For Each wsA In wbA.Worksheets
        With wsA
            lR = .Cells.Find("*", after:=Cells(Rows.Count, 1), _
                searchorder:=xlRows, searchdirection:=xlPrevious).Row
            lC = .Cells.Find("*", after:=Cells(1, Columns.Count), _
                searchorder:=xlColumns, searchdirection:=xlPrevious).Column
'           .PrintPreview
        End With
            
        With wsA.PageSetup
            .PrintArea = ActiveSheet.Range(Cells(1, 1), Cells(lR, lC)).Address
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
    Next wsA

    Debug.Print myFile

    '-----Export to PDF if a folder was selected------
    If myFile <> "False" Then
        wbA.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=myFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If

    ActiveWorkbook.FollowHyperlink Address:=strPath, NewWindow:=True

End Sub



'Old Code

''    Application.PrintCommunication = False
''              LastRow = Range("A:N").SpecialCells(xlCellTypeLastCell).Row
'                LastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
'                LastRow = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Row - 1
'              wsA.Activate
'              ActiveSheet.UsedRange
'              ActiveSheet.ResetAllPageBreaks

''              ActiveSheet.Cells.ClearFormats
'            .PrintArea = "$A$2:$R$" & LastRow
'            .PrintArea = LastRow
'    Application.PrintCommunication = True




' For Each wsA In wbA.Worksheets
'        With wsA
'            lR = .Cells.Find("*", after:=Cells(Rows.Count, 1), _
'                searchorder:=xlRows, searchdirection:=xlPrevious).Row
'            lC = .Cells.Find("*", after:=Cells(1, Columns.Count), _
'                searchorder:=xlColumns, searchdirection:=xlPrevious).Column
''           .PageSetup.PrintArea = .Range(Cells(1, 1), Cells(lR, lC)).Address
''           .PageSetup.Orientation = xlLandscape
'            'to immediately go to printout use: (else comment out)
''           .PrintOut Copies:=1, Collate:=True, _
'            IgnorePrintAreas:=False
'            'to see preview first use: (else comment out)
''           .PrintPreview
'        End With
