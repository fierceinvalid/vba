Sub SaveMessageAsPDF()
     
    Dim Selection As Selection
    Dim obj As Object
    Dim Item As MailItem
     
 
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    Set Selection = Application.ActiveExplorer.Selection
    
     Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
'
'    Dim objRegExp As Object
'    Set objRegExp = CreateObject("VBScript.RegExp")
'    objRegExp.Pattern = "Wings ID#\s\d{1,10}"
'    objRegExp.IgnoreCase = True
'    objRegExp.Global = True

     Dim regex As Object
     Dim matchCollection As Object
     Dim extractedString As String
     Dim str As String
  '   Set regex = CreateObject("VBScript.RegExp")
     Set regex = New RegExp

           With regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .MultiLine = True
            .Pattern = "are"
        End With

    Set matchCollection = regex.Execute(Search_Email)
    If matchCollection.Count <> 0 Then
    extractedString = matchCollection.Item(0).SubMatches.Item(0)
End If

'Debug.Print matchCollection.Item(0)

For Each obj In Selection
 
    Set Item = obj
    
    Dim FSO As Object, TmpFolder As Object
    Dim sName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set tmpFileName = FSO.GetSpecialFolder(2)
    
 '   sName = Item.Subject
'    sName = Item.To
    sName = extractedString + "1"
'    ReplaceCharsForFileName sName, "-"
    tmpFileName = tmpFileName & "\" & sName & ".mht"
    
    Item.SaveAs tmpFileName, olMHTML
    
    
Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpFileName, Visible:=True)
  
    Dim WshShell As Object
    Dim SpecialPath As String
    Dim strToSaveAs As String
    Set WshShell = CreateObject("WScript.Shell")
'    MyDocs = WshShell.SpecialFolders(16)
                                             MyDocs = "file path"
       
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
