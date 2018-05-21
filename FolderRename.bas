Sub RenameSubfolders()

Dim FileSystem As Object
Dim Folder As Object
Dim SubFolder As Object
Dim InitialPath As String
Dim sOld As String
Dim sNew As String
Dim answer As Integer

Dim Folder1 As Object
Dim SubFolder1 As Object
Dim InitialPath1 As String
Dim oldFolder1 As String
Dim newFolder1 As String
Dim sCommName As String
    
    sCommName = Range("F2")
    sDir = "C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans"
    
    InitialPath = sDir + "\" + sCommName + "\"
    

    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)

    For Each SubFolder In Folder.subfolders
        sOld = SubFolder.Name
        If InStr(1, sOld, "1") > 0 Then
            If sOld = "1 Details" Then
            'do nothing
        Else
            
            answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & sOld & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & "1 Details", vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                sNew = Replace(sOld, sOld, "1 Details")
            Name InitialPath & sOld As InitialPath & sNew
            Else
                'do nothing
            End If
        End If

        ElseIf InStr(1, sOld, "2") > 0 Then
         If sOld = "2 Documentation" Then
           'do nothing
        Else
            sNew = Replace(sOld, sOld, "2 Documentation")
            Name InitialPath & sOld As InitialPath & sNew
        End If
        ElseIf InStr(1, sOld, "3") > 0 Then
         If sOld = "3 Business Member Financials" Then
           'do nothing
        Else
            sNew = Replace(sOld, sOld, "3 Business Member Financials")
            Name InitialPath & sOld As InitialPath & sNew
        End If
        ElseIf InStr(1, sOld, "4") > 0 Then
        If sOld = "4 Guarantor Financials" Then
          'do nothing
        Else
            sNew = Replace(sOld, sOld, "4 Guarantor Financials")
            Name InitialPath & sOld As InitialPath & sNew
        End If
        ElseIf InStr(1, sOld, "5") > 0 Then
        If sOld = "5 Collateral" Then
           'do nothing
        Else
            sNew = Replace(sOld, sOld, "5 Collateral")
            Name InitialPath & sOld As InitialPath & sNew
        End If
        ElseIf InStr(1, sOld, "6") > 0 Then
        If sOld = "6 Miscellaneous" Then
            'do nothing
        Else
            sNew = Replace(sOld, sOld, "6 Miscellaneous")
            Name InitialPath & sOld As InitialPath & sNew
        End If
        End If
'
'        If InStr(1, sOld, "#") > 0 Then
'            sNew = Replace(sOld, "#", Format(Month("01/" & Right(sOld, 3)), "00"))
'            Name InitialPath & sOld As InitialPath & sNew
'        End If

    Next SubFolder


InitialPath1 = "C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans\123456 Derek Industries\1 Details\"

    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder1 = FileSystem.GetFolder(InitialPath1)

    For Each SubFolder1 In Folder1.subfolders
        oldFolder1 = SubFolder1.Name

         If InStr(1, oldFolder1, "10") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "10 Document Checklist")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
        ElseIf InStr(1, oldFolder1, "11") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "11 Commercial Loan Presentation")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
        ElseIf InStr(1, oldFolder1, "12") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "12 File Comments")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
       End If
    Next SubFolder1
    
    
    
    InitialPath1 = "C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans\123456 Derek Industries\2 Documentation\"

    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder1 = FileSystem.GetFolder(InitialPath1)

    For Each SubFolder1 In Folder1.subfolders
        oldFolder1 = SubFolder1.Name

         If InStr(1, oldFolder1, "20") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "20 Document Checklist")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
        ElseIf InStr(1, oldFolder1, "21") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "21 Commercial Loan Presentation")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
        ElseIf InStr(1, oldFolder1, "22") > 0 Then
            newFolder1 = Replace(oldFolder1, oldFolder1, "22 File Comments")
            Name InitialPath1 & oldFolder1 As InitialPath1 & newFolder1
       End If
    Next SubFolder1


End Sub



'InStr(4, "Excel-Trick", "c") will result into 10 as here we are starting the search for “c”
'from the 4th character and hence Instr gives us the position of second “c” (Excel-Trick) in the ‘parent_string’.
