Sub FolderFileRenamer()

'--File Stuff ---
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
Dim xString7 As String
Dim spltFilename As String
Dim spltNew As String
Dim OldFile As String
Dim NewFile As String
Dim lasDigit As String
Dim i As Byte
Dim dump As VbMsgBoxResult
Dim answer As Integer
'Dim f As Object
Dim sPath2 As String


'---Folder Stuff----

Dim FileSystem As Object
Dim Folder As Object
Dim SubFolder As Object
Dim InitialPath As String
Dim sOld As String
Dim sNew As String
'Dim answer As Integer

'Dim Folder1 As Object
'Dim SubFolder1 As Object
'Dim InitialPath1 As String
Dim oldFolder As String
Dim newFolder As String
Dim sCommName As String
Dim Ret
Dim diaFolder As FileDialog
Dim Fname As String
Dim sBrowsePath As String
Dim result As String
Dim lasFolder As String

Dim var1 As String
Dim var10 As String
Dim var11 As String
Dim var12 As String
Dim var13 As String
Dim var14 As String
Dim var15 As String
Dim var16 As String
Dim var17 As String

Dim var2 As String
Dim var20 As String
Dim var21 As String
Dim var22 As String
Dim var23 As String
Dim var24 As String
Dim var25 As String

Dim var3 As String
Dim var30 As String
Dim var31 As String
Dim var32 As String
Dim var33 As String
Dim var34 As String
Dim var35 As String
Dim var36 As String
Dim var37 As String

Dim var4 As String
Dim var40 As String
Dim var41 As String

Dim var5 As String
Dim var50 As String
Dim var51 As String
Dim var52 As String
Dim var53 As String
Dim var54 As String
Dim var55 As String
Dim var56 As String
Dim var57 As String
Dim var58 As String
Dim var59 As String

Dim var6 As String
Dim var60 As String
Dim var61 As String
Dim var62 As String
Dim var63 As String
Dim sCommPath As String



'--Folders under 1 Details----
var1 = "1 Details"
var10 = "10 Document Checklist"
var11 = "11 Commercial Loan Presentation"
var12 = "12 File Comments"
var13 = "13 Business Loan and Deposit Application"
var14 = "14 Personal Loan and Deposit Application"
var15 = "15 Proposals and Commitments"
var16 = "16 Corporate Documentation"
var17 = "17 Misc Corporate Information"

'--Folders under 2----
var2 = "2 Documentation"
var20 = "20 Note"
var21 = "21 Loan Agreement"
var22 = "22 Mortgages & Security Agreements"
var23 = "23 Guaranty"
var24 = "24 Misc Signed Agreements by Bank & Borrower"
var25 = "25 Misc Note Specific Documents"

'--Folders under 3----
var3 = "3 Business Member Financials"
var30 = "30 Business Financial Statements"
var31 = "31 Business Interim Statements"
var32 = "32 Business Tax Returns"
var33 = "33 A&R Listings and Borrowing Base Certificates"
var34 = "34 Misc Financial Information"
var35 = "35 Financial Spreads"
var36 = "36 - Rent Rolls"
var37 = "37 - Leases"

'--Folders under 4----
var4 = "4 Guarantor Financials"
var40 = "40 Personal Financial Information"
var41 = "41 - Background Checks"

'--Folders under 5----
var5 = "5 Collateral"
var50 = "50 Insurance Binders"
var51 = "51 UCC Search & Filings"
var52 = "52 Lien Cards"
var53 = "53 Misc Collateral Control Documents"
var54 = "54 Title Work"
var55 = "55 Collateral Valuation Documentation"
var56 = "56 Collateral Review & Condition Documentation"
var57 = "57 Site Visit"
var58 = "58 Flood"
var59 = "59 Construction Draws"

'--Folders under 6----
var6 = "6 Miscellaneous"
var60 = "60 Correspondence"
var61 = "61 Escrow Analysis"
var62 = "62 Misc Information"
var63 = "63 Paid Loan File"

sCommPath = "file path of main folder"

    '~~> Specify your start folder here
    result = BrowseForFolder("file path of main folder")
    Select Case result
        Case Is = False
            result = "an invalid folder!"
        Case Else
            'don't change anything
    End Select

If result = sCommPath Then
    Exit Sub
End If
If Len(result) = 0 Then
Exit Sub
End If

'--Main Folder------
    InitialPath = result + "\"
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)

    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

        If Left(Trim(oldFolder), 1) = 1 Then
'         If InStr(1, oldFolder, "1") > 0 Then
            If oldFolder = var1 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var1, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var1)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 1) = 2 Then
            If oldFolder = var2 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var2, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var2)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 1) = 3 Then
             If oldFolder = var3 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var3, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var3)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 1) = 4 Then
            If oldFolder = var4 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var4, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var4)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
             ElseIf Left(Trim(oldFolder), 1) = 5 Then
             If oldFolder = var5 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var5, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var5)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
         ElseIf Left(Trim(oldFolder), 1) = 6 Then
            If oldFolder = var6 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var6, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var6)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder


'--1 Details Folder------
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + "1 Details\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

        If Left(Trim(oldFolder), 2) = 10 Then
            If oldFolder = var10 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var10, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var10)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 11 Then
            If oldFolder = var11 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var11, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var11)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 12 Then
            If oldFolder = var12 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var12, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var12)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 13 Then
            If oldFolder = var13 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var13, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var13)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 14 Then
            If oldFolder = var14 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var14, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var14)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 15 Then
            If oldFolder = var15 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var15, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var15)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 16 Then
             If oldFolder = var16 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var16, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var16)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
      ElseIf Left(Trim(oldFolder), 2) = 17 Then
            If oldFolder = var17 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var12, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var12)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
  Else
        'do nothing
End If

'--2 Documentation------
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + var2 + "\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

        If Left(Trim(oldFolder), 1) = 20 Then
          If oldFolder = var20 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var20, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var20)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 21 Then
            If oldFolder = var21 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var21, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var21)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 22 Then
            If oldFolder = var22 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var22, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var22)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 23 Then
            If oldFolder = var23 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var23, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var23)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 24 Then
            If oldFolder = var24 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var24, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var24)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 25 Then
            If oldFolder = var25 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var25, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var25)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
   Else
        'do nothing
End If
  
'--3 Business Member Financials------
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + var3 + "\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

         If Left(Trim(oldFolder), 1) = 30 Then
                 If oldFolder = var30 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var30, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var30)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 31 Then
                 If oldFolder = var31 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var31, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var31)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 32 Then
                 If oldFolder = var32 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var32, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var32)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 33 Then
                  If oldFolder = var33 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var33, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var33)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 34 Then
               If oldFolder = var34 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var34, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var34)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 35 Then
                 If oldFolder = var35 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var35, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var35)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 36 Then
                 If oldFolder = var36 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var36, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var36)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 37 Then
                 If oldFolder = var37 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var37, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var37)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
  Else
        'do nothing
End If

    '--4 Guarantor Financials------
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + var4 + "\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

         If Left(Trim(oldFolder), 1) = 40 Then
                 If oldFolder = var40 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var40, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var40)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 41 Then
               If oldFolder = var41 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var41, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var41)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
  Else
        'do nothing
End If

'--5 Collateral------
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + var5 + "\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

         If Left(Trim(oldFolder), 1) = 50 Then
                If oldFolder = var50 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var50, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var50)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 51 Then
                 If oldFolder = var51 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var51, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var51)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 52 Then
                If oldFolder = var52 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var52, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var52)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
      ElseIf Left(Trim(oldFolder), 2) = 53 Then
                If oldFolder = var53 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var53, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var53)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 54 Then
                If oldFolder = var54 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var54, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var54)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 55 Then
                 If oldFolder = var55 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var55, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var55)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
     ElseIf Left(Trim(oldFolder), 2) = 56 Then
                If oldFolder = var56 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var56, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var56)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 57 Then
                 If oldFolder = var57 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var57, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var57)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
      ElseIf Left(Trim(oldFolder), 2) = 58 Then
                 If oldFolder = var58 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var58, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var58)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 59 Then
                 If oldFolder = var59 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var59, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var59)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
  Else
        'do nothing
End If

'--6 Miscellaneous-----
    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
    InitialPath = sCommPath + "\" + lasFolder + "\" + var6 + "\"
If Dir(InitialPath, vbDirectory) <> vbNullString Then
    Set FileSystem = CreateObject("Scripting.filesystemobject")
    Set Folder = FileSystem.GetFolder(InitialPath)
    Debug.Print InitialPath
    For Each SubFolder In Folder.subfolders
        oldFolder = SubFolder.Name

        If Left(Trim(oldFolder), 1) = 60 Then
                If oldFolder = var60 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var60, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var60)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
      ElseIf Left(Trim(oldFolder), 2) = 61 Then
                If oldFolder = var61 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var61, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var61)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       ElseIf Left(Trim(oldFolder), 2) = 62 Then
                If oldFolder = var62 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var62, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var62)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
        ElseIf Left(Trim(oldFolder), 2) = 63 Then
                If oldFolder = var63 Then 'do nothing
            Else: answer = MsgBox("CHANGE FOLDER NAME OF:" & vbNewLine & vbNewLine & oldFolder & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & var63, vbYesNo + vbQuestion, "Empty Sheet")
                If answer = vbYes Then
                    newFolder = Replace(oldFolder, oldFolder, var63)
                    Name InitialPath & oldFolder As InitialPath & newFolder
                Else 'do nothing
                End If
            End If
       End If
    Next SubFolder
  Else
        'do nothing
End If










    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
'    objRegExp.Pattern = "\d(\s|\w|-)"
    objRegExp.Pattern = ".pdf"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True

    Dim colFiles As Collection
    Set colFiles = New Collection
    
    sPath2 = sCommPath + "\" + lasFolder + "\"

    RecursiveFileSearch sPath2, objRegExp, colFiles, objFSO

    For Each f In colFiles
    'Debug.Print (f)
        'Insert code here to do something with the matched files
        
        
        
 '---Match and change all files in format: YYYY-.M-.DD   ----
        
        
        Dim regex As Object, str As String
        Set regex = CreateObject("VBScript.RegExp")
 
            With regex
                .Pattern = "^\d{4}(\.|-)\d{1}(\.|-)\d{2}"
                .Global = True
            End With
     

        xString = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        Set matches = regex.Execute(xString)
        For Each Match In matches
        
        spltSpace = Left(xString, 9)
        lasDigit = Mid(spltSpace, 1)
        
         
        sDateFinal = Replace(spltSpace, "-", ".")
        spltYear = Split(sDateFinal, ".")(0)
        spltMonth = Split(sDateFinal, ".")(1)
        spltDay = Split(sDateFinal, ".")(2)
        
        spltFilename = Mid(xString, 2 + Len(xString) - InStr(StrReverse(xString), lasDigit))
 
        
        strName = spltYear + "." + "0" + spltMonth + "." + spltDay + "" + Mid(xString, 10)
        g = Replace(f, xString, strName)
     '   strFile = g & ".pdf"
        strFile = g
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
    
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
    
        

        Next Match
        




 '---Match and change all files in format: YYYY-.MM-.D(Space)   ----
        
        
        Dim regex4 As Object, str4 As String
        Set regex4 = CreateObject("VBScript.RegExp")
 
            With regex4
                .Pattern = "^\d{4}(\.|-)\d{2}(\.|-)\d{1}\s"
                .Global = True
            End With
     

        xString = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        Set matches = regex4.Execute(xString)
        For Each Match In matches
        
        spltSpace = Left(xString, 9)
        lasDigit = Mid(spltSpace, 1)
        
         
        sDateFinal = Replace(spltSpace, "-", ".")
        spltYear = Split(sDateFinal, ".")(0)
        spltMonth = Split(sDateFinal, ".")(1)
        spltDay = Split(sDateFinal, ".")(2)
        
        spltFilename = Mid(xString, 2 + Len(xString) - InStr(StrReverse(xString), lasDigit))
 
        
        strName = spltYear + "." + spltMonth + "." + "0" + spltDay + " " + Mid(xString, 10)
        g = Replace(f, xString, strName)
     '   strFile = g & ".pdf"
        strFile = g
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
  
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
    
        

        Next Match
        
        
        
        
         '---Match and change all files in format: YYYY-.M-.D(Space)   ----
        
        
        Dim regex5 As Object, str5 As String
        Set regex5 = CreateObject("VBScript.RegExp")
 
            With regex5
                .Pattern = "^\d{4}(\.|-)\d{1}(\.|-)\d{1}\s"
                .Global = True
            End With
     

        xString = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        Set matches = regex5.Execute(xString)
        For Each Match In matches
        
        spltSpace = Left(xString, 8)
        lasDigit = Mid(spltSpace, 1)
        
         
        sDateFinal = Replace(spltSpace, "-", ".")
        spltYear = Split(sDateFinal, ".")(0)
        spltMonth = Split(sDateFinal, ".")(1)
        spltDay = Split(sDateFinal, ".")(2)
        
        spltFilename = Mid(xString, 2 + Len(xString) - InStr(StrReverse(xString), lasDigit))
 
        
        strName = spltYear + "." + "0" + spltMonth + "." + "0" + spltDay + " " + Mid(xString, 9)
        g = Replace(f, xString, strName)
        strFile = g
        'strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
   
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
    
        

        Next Match
        
        
   '---Match and change all files in format: YYYY   ----
   
   
        Dim regex2 As Object, str2 As String
        Set regex2 = CreateObject("VBScript.RegExp")
 
            With regex2
                .Pattern = "^\d{4}\s"
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
        
           strName = strYearDM + "" + spltFilename
        g = Replace(f, xString2, strName)
        strFile = g
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
      
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
        
        
         '---Match and change all files in format: YYYY   ----
   
   
        Dim regex6 As Object, str6 As String
        Set regex6 = CreateObject("VBScript.RegExp")
 
            With regex6
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
        strFile = g
        'strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
     
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
        
        
        
  '---Match and change all files in format: YYYY(Letter)   ----
        
        
            Dim regex3 As Object, str3 As String
        Set regex3 = CreateObject("VBScript.RegExp")
 
            With regex3
                .Pattern = "^\d{4}[A-Za-z]"
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
        strFile = g
        'strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
      
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
              



'---Match and change all files in format: YYYY.MM(Letter)   ----
        
        
            Dim regex7 As Object, str7 As String
        Set regex7 = CreateObject("VBScript.RegExp")
 
            With regex7
                .Pattern = "^\d{4}(\.|-)\d{2}\s"
                .Global = True
            End With
            
                

        xString7 = Mid(f, 2 + Len(f) - InStr(StrReverse(f), "\"))
        
        
        Set matches = regex7.Execute(xString3)
        For Each Match In matches
        
        
        
        
        spltSpace = Left(xString7, 7)
        spltFilename = Right(xString7, Len(xString7) - Len(spltSpace))
        'spltFilename = Split(xString2, "8", 2)(1)
        
        
         
        
        strYearDM = spltSpace + ".31"
        
        'h = Replace(f, strYearDM, xString2)
        
        'Debug.Print Match.Value
        'Debug.Print strYearDM
        
           strName = strYearDM + "" + spltFilename
        g = Replace(f, xString7, strName)
        strFile = g
        'strFile = g & ".pdf"
        myFile = strPath & strFile
        OldFile = f
        NewFile = strFile
        
      
        answer = MsgBox("CHANGE FILE NAME OF:" & vbNewLine & vbNewLine & xString & vbNewLine & vbNewLine & "TO:" & vbNewLine & vbNewLine & strName, vbYesNo + vbQuestion, "Empty Sheet")
            If answer = vbYes Then
                Name OldFile As NewFile
            Else
                'do nothing
            End If
              
        '   Name OldFile As NewFile
        
        Debug.Print g
        

        Next Match
        
        
    Next

    'Garbage Collection
    Set objFSO = Nothing
    Set objRegExp = Nothing
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







'
'
''--5 Collateral------
'    lasFolder = Mid(result, 2 + Len(result) - InStr(StrReverse(result), "\"))
'    InitialPath = "C:\Users\1933\Desktop\Comm Loans Testing\Full Files" + "\" + lasFolder + "\" + var5 + "\"
'If Dir(InitialPath, vbDirectory) <> vbNullString Then
'    Set FileSystem = CreateObject("Scripting.filesystemobject")
'    Set Folder = FileSystem.GetFolder(InitialPath)
'    Debug.Print InitialPath
'    For Each SubFolder In Folder.subfolders
'        oldFolder = SubFolder.Name
'
'         If InStr(1, oldFolder, "50") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var50)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "51") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var51)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "52") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var52)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "53") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var53)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "54") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var54)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "55") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var55)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "56") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var56)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "57") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var57)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "58") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var58)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'        ElseIf InStr(1, oldFolder, "59") > 0 Then
'            newFolder = Replace(oldFolder, oldFolder, var59)
'            Name InitialPath & oldFolder As InitialPath & newFolder
'       End If
'    Next SubFolder
'  Else
'        'do nothing
'End If





Function RecursiveFileSearch(ByVal targetFolder As String, ByRef objRegExp As Object, _
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
    Set objSubFolders = objFolder.subfolders
    For Each objSubfolder In objSubFolders
        RecursiveFileSearch objSubfolder, objRegExp, matchedFiles, objFSO
    Next

    'Garbage Collection
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objSubFolders = Nothing

End Function
