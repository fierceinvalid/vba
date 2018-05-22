ub RenameSubfolders()

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
Dim Ret
Dim diaFolder As FileDialog
Dim Fname As String
Dim sBrowsePath As String
Dim result As String

    '~~> Specify your start folder here
    result = BrowseForFolder("C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans")
    Select Case result
        Case Is = False
            result = "an invalid folder!"
        Case Else
            'don't change anything
    End Select
'    MsgBox "You selected " & result, _
'        vbOKOnly + vbInformation
'

'''Open the file dialog
''On Error GoTo ErrorHandler
'Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
'diaFolder.AllowMultiSelect = False
'diaFolder.Title = "Select a folder then hit OK"
'diaFolder.Show
'Fname = diaFolder.SelectedItems(1)
'Debug.Print Fname


'ErrorHandler:
'Msg = "No folder selected, you must select a folder for program to run"
'Style = vbError
'Title = "Need to Select Folder"
'Response = MsgBox(Msg, Style, Title)

'    '~~> Specify your start folder here
'    Ret = BrowseForFolder("C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans")
'
'    sCommName = Range("F2")
'
'     If Not IsEmpty(Range("F2")) Then
'        sCommName = Range("F2")
'     Else:
'        MsgBox "Please type in the Folder name of the Commercial Loan you want to check"
'    Exit Sub
'    End If
    
'    sDir = "C:\Users\dbalk\Desktop\Work\Eforms\Commercial Loans"
'    BrowseForFolder = sBrowsePath
    InitialPath = result + "\"
'    InitialPath = sDir + "\" + sCommName + "\"
    
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

MsgBox ("Folders and Files Have been Checked")
End Sub



'InStr(4, "Excel-Trick", "c") will result into 10 as here we are starting the search for “c”
'from the 4th character and hence Instr gives us the position of second “c” (Excel-Trick) in the ‘parent_string’.


Function BrowseForFolder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level

    Dim ShellApp As Object

     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a Commercial Loan", 0, OpenAt)

     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0

     'Destroy the Shell Application
    Set ShellApp = Nothing

     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
'    Select Case Mid(BrowseForFolder, 2, 1)
'    Case Is = ":"
'        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
'    Case Is = "\"
'        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
'    Case Else
'        GoTo Invalid
'    End Select

    Exit Function

Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function



