Option Explicit

Sub AcctMain()
    
Dim sDocTypeNum As String
Dim sDocTypeName As String
Dim Rng As Range



    'Look for number in cell C1, split it by dash, display message if not there
    If Not IsEmpty(Range("C1")) Then
        sDocTypeNum = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
     Else
        InputBox ("No Document Number Found in Cell C1. Please Enter A DocType, Number and Macro to run")
    End If
    
    'Use found doc number to match number in column, then run correct module.
    If Trim(sDocTypeNum) <> "" Then
    With Workbooks("personal.xlsm").Sheets("personal").Range("B:B") 'searches all of column B WorkBookSinglePDFStart
        Set Rng = .Find(What:=sDocTypeNum, _
                        After:=.Cells(.Cells.Count), _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        SearchOrder:=xlByRows, _
                        SearchDirection:=xlNext, _
                        MatchCase:=False)
        If Not Rng Is Nothing Then
             Call WorkBookSinglePDFStart
        
        ElseIf Trim(sDocTypeNum) <> "" Then
        With Workbooks("personal.xlsm").Sheets("personal").Range("F:F") 'searches all of column  F BatchExportWbStart
            Set Rng = .Find(What:=sDocTypeNum, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                Call BatchExportWbStart
            
            ElseIf Trim(sDocTypeNum) <> "" Then
            With Workbooks("personal.xlsm").Sheets("personal").Range("H:H") 'searches all of column H Wells
                Set Rng = .Find(What:=sDocTypeNum, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
                If Not Rng Is Nothing Then
                    Call WellsFargoBankStart
                Else
                    InputBox (sDocTypeName + " " + sDocTypeNum + " " + "does not have a Macro that is associated with it. What Macro would you like to run")
                End If
            End With
            End If
        End With
        End If
    End With
    End If
    
End Sub




'old code
' Put this into BatchExportWb
'    If IsEmpty(Range("C1")) And IsEmpty(Range("B1")) Then
'       InputBox ("No Document Number Found. Please Enter the Name & Number of the Document Below")
'     ElseIf IsEmpty(Range("C1")) Then
'        sDocTypeNum = LTrim(RTrim(Split(Range("B1").Value, "-")(0)))
'    ElseIf IsEmpty(Range("B1")) Then
'        sDocTypeNum = LTrim(RTrim(Split(Range("C1").Value, "-")(0)))
'    End If



    
'    Select Case (sDocTypeName)
'        Case "Garda", "Returned Items", "Unposted Items"
'            Call WorkBookSinglePDFStart
'        Case "ATM Deposits In Transit"
'           Call BatchExportWbStart
'        Case "Wells Fargo Bank"
'        Case Workbooks("personal.xlsm").Sheets("personal").Range("J6").Value
'       Case Workbooks("personal.xlsm").Sheets("personal").Range(Cells(1, 1), Cells(TotalRows, 1)).Value
'           Call WellsFargoBankStart
'        Case Else
'            Select Case MsgBox(Prompt:=sDocType + " " + "does not have a Macro that is associated with it. What Macro would you like to run", _
'                       Buttons:=vbYesNoCancel)
'                Case vbYes
'                    Call WorkBookSinglePDFStart
'                Case vbNo
'                     Call WellsFargoBankStart
'                Case vbCancel
'
'            End Select
'    End Select


'Sub Macro1()
'Dim sDocType As String
'
'    Select Case MsgBox(Prompt:=sDocTypeName + " " + "does not have a Macro that is associated with it. What Macro would you like to run", _
'                       Buttons:=vbYesNoCancel)
'        Case vbYes
'            Range("A1") = Range("A1") + 1 'increment hidden sequence num
'            nosoumission = noclient & Range("A1")
'            Sheets("Sheet1").Range("G9") = nosoumission
'            Range("B13").ClearContents
'            Application.Dialogs(xlDialogSaveAs).Show
'        Case vbNo
'            Range("A1") = Range("A1") + 1 'increment hidden sequence num
'            nosoumission = noclient & Range("A1")
'            Sheets("Sheet1").Range("G9") = nosoumission
'            Application.Dialogs(xlDialogSaveAs).Show
'        Case vbCancel
'            Exit Sub
'    End Select
'End Sub
