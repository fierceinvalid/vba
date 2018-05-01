Option Explicit
'Dim sDocTypeNum As String
'Dim sDocTypeName As String
'Dim sDocTypeNumFinal As String

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


'    sNumber = LTrim(RTrim(Split(Range("D1").Value, "-")(0)))
'    sDocType = LTrim(RTrim(Split(Range("D1").Value, "-")(1)))
'
    sDateFinal = Replace(sDate, "/", ".")
'    sNumberFinal = Replace(sNumber, ".", "")
    
    sDirPath = "C:\Users\dbalk\Desktop\Work\Eforms\Accounting\GL Reconciliation"
    strPath = sDirPath & "\" & sDocTypeName
    

    If Dir("C:\Users\dbalk\Desktop\Work\Eforms\Accounting\GL Reconciliation\" + sDocTypeName, vbDirectory) = "" Then
        MkDir Path:="C:\Users\dbalk\Desktop\Work\Eforms\Accounting\GL Reconciliation\" + sDocTypeName
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
    
    '------Find Latest Date in Month------
'    If spltMonthNew = "01" Then
'        spltDayEnd = "31"
'    ElseIf spltMonthNew = "02" Then
'        spltDayEnd = "28"
'    ElseIf spltMonthNew = "03" Then
'        spltDayEnd = "31"
'    ElseIf spltMonthNew = "04" Then
'        spltDayEnd = "30"
'     ElseIf spltMonthNew = "05" Then
'        spltDayEnd = "31"
'     ElseIf spltMonthNew = "06" Then
'        spltDayEnd = "30"
'     ElseIf spltMonthNew = "07" Then
'        spltDayEnd = "31"
'     ElseIf spltMonthNew = "08" Then
'        spltDayEnd = "31"
'     ElseIf spltMonthNew = "09" Then
'        spltDayEnd = "30"
'     ElseIf spltMonthNew = "10" Then
'        spltDayEnd = "31"
'     ElseIf spltMonthNew = "11" Then
'        spltDayEnd = "30"
'     ElseIf spltMonthNew = "12" Then
'        spltDayEnd = "31"
'    End If
    
    
    If Len(sNumberFinal) > 10 Then
        sNumberFinal = "MULTIPLE"
    End If
    
    
    
    
    'add day and month back together
    'use date on worksheet
    'strName = spltMonthNew + "." + spltDayNew + "." + spltYearNew + " " + sNumberFinal + " " + sDocType
    
    'use end of month day in format YYYY.mm.dd + No DocType
  '  strName = spltYearNew + "." + spltMonthNew + "." + spltDayNew + " " + sNumberFinal + " " + sDocType
    
    strName = spltYearNew + "." + spltMonthNew + "." + spltDayNew + " " + sDocTypeNumFinal + " " + sDocTypeName
    
    
    
    'replace spaces and periods in sheet name
    'strNameNew = Replace(strName, " ", ".")
    strNameNew = strName

    'create default name for savng file
    strFile = strNameNew & ".pdf"
    myFile = strPath & strFile
    
    
          Application.PrintCommunication = False
          For Each wsA In wbA.Worksheets
              wsA.Activate
              ActiveSheet.UsedRange
              With wsA.PageSetup
                    .PrintArea = ""
                    .Orientation = xlLandscape
                    .Zoom = False
                    '.PrintArea = Worksheets(ReportWsName).UsedRange
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                End With
            Next wsA
            Application.PrintCommunication = True
    
'    For Each wsA In wbA.Worksheets
'        wsA.PageSetup.Orientation = xlLandscape
'
'
'    Next
'

    Debug.Print myFile

    'export to PDF if a folder was selected
        If myFile <> "False" Then
        ActiveWorkbook.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=myFile, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
  
    End If

 ActiveWorkbook.FollowHyperlink Address:=strPath, NewWindow:=True

End Sub
