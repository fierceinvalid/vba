Sub SaveMessageAsPDF()
     
    Dim Selection As Selection
    Dim obj As Object
    Dim Item As MailItem
    Dim regex As Object
    Dim matchCollection As Object
    Dim extractedString As String
    Dim str As String
    Dim mailContent As String
    Dim sMail As String
    Dim mail As Outlook.MailItem
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    Set Selection = Application.ActiveExplorer.Selection
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim Email As Outlook.MailItem
    Dim Matches As Variant
    Dim RegExp As Object
    Dim Pattern As String

'   Const Email As String = "RE: Blah blah - Blah blah, Wings# 12345678, Blah blah blah"

'
'     Set regex = CreateObject("vbscript.regexp")
'
'
'           With regex
'            .Global = False     'Check the first instance
'            .MultiLine = True
'            .IgnoreCase = False
'            .MultiLine = True
'            .Pattern = "Wings# \d{3,9}"
'        End With
'
'
'    Set matchCollection = regex.Execute(sMail)
'    If matchCollection.Count <> 0 Then
'    extractedString = matchCollection.Item(0)
'End If

'Debug.Print matchCollection.Item(0)

Set RegExp = CreateObject("VbScript.RegExp")

    If TypeOf Item Is Outlook.MailItem Then

        Pattern = "Wings# \d{3,9}"
        With RegExp
            .Global = False
            .Pattern = Pattern
            .IgnoreCase = True
             Set Matches = .Execute(Item.Body)
        End With

 If Matches.Count > 0 Then
            Debug.Print Item.Subject ' Print on Immediate Window
'            Set Email = Item.Forward
'                Email.Subject = Item.Subject
'                Email.Recipients.Add "0m3r@Email.com"
'                Email.Save
'                Email.Send

        End If
    End If




For Each obj In Selection

    Set Item = obj
    
    Dim FSO As Object, TmpFolder As Object
    Dim sName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set tmpFileName = FSO.GetSpecialFolder(2)
    
 '   sName = Item.Subject
'    sName = Item.To
    sName = extractedString
'    ReplaceCharsForFileName sName, "-"
    tmpFileName = tmpFileName & "\" & sName & ".mht"
    
    Item.SaveAs tmpFileName, olMHTML
    
    
Set wrdDoc = wrdApp.Documents.Open(Filename:=tmpFileName, Visible:=True)
  
    Dim WshShell As Object
    Dim SpecialPath As String
    Dim strToSaveAs As String
    Set WshShell = CreateObject("WScript.Shell")
'    MyDocs = WshShell.SpecialFolders(16)
    MyDocs = "C:\Users\dbalk\Desktop\Work\Eforms\Payment Systems\Debit Fraud Disputes"
       
strToSaveAs = MyDocs & "\" & sName & ".pdf"
 
' check for duplicate filenames
' if matched, add the current time to the file name
If FSO.FileExists(strToSaveAs) Then
   sName = sName & Format(Now, "hhmmss")
   strToSaveAs = MyDocs & "\" & sName & ".pdf"
End If
  
wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
    strToSaveAs, ExportFormat:=wdExportFormatPDF, _
    OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
    Range:=wdExportAllDocument, From:=0, To:=0, Item:= _
    wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
    CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
    BitmapMissingFonts:=True, UseISO19005_1:=False
             
    

Next obj
    wrdDoc.Close
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    Set WshShell = Nothing
    Set obj = Nothing
    Set Selection = Nothing
    Set Item = Nothing
 
End Sub
 
' This function removes invalid and other characters from file names
Private Sub ReplaceCharsForFileName(sName As String, sChr As String)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
  sName = Replace(sName, "&", sChr)
  sName = Replace(sName, "%", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, " ", sChr)
  sName = Replace(sName, "{", sChr)
  sName = Replace(sName, "[", sChr)
  sName = Replace(sName, "]", sChr)
  sName = Replace(sName, "}", sChr)
  sName = Replace(sName, "!", sChr)
End Sub

