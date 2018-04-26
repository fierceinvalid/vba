Attribute VB_Name = "SaveAllEmail"
Sub SaveMessageAsPDF()
     
    Dim Selection As Selection
    Dim obj As Object
    Dim Item As MailItem
     
 
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    Set Selection = Application.ActiveExplorer.Selection

For Each obj In Selection
 
    Set Item = obj
    
    Dim FSO As Object, TmpFolder As Object
    Dim sName As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set tmpFileName = FSO.GetSpecialFolder(2)
    
    sName = Item.To
    ReplaceCharsForFileName sName, "-"
    tmpFileName = tmpFileName & "\" & sName & ".mht"
    
    Item.SaveAs tmpFileName, olMHTML
    
    
Set wrdDoc = wrdApp.Documents.Open(FileName:=tmpFileName, Visible:=True)
  
    Dim WshShell As Object
    Dim SpecialPath As String
    Dim strToSaveAs As String
    Set WshShell = CreateObject("WScript.Shell")
    MyDocs = "\\prshsan02\apps\Apps\Eforms\Deposit Services\Correspondence\Duplicate Mobile Deposit"
       
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
'  sName = Replace(sName, "/", sChr)
'  sName = Replace(sName, "\", sChr)
'  sName = Replace(sName, ":", sChr)
'  sName = Replace(sName, "?", sChr)
'  sName = Replace(sName, Chr(34), sChr)
'  sName = Replace(sName, "<", sChr)
'  sName = Replace(sName, ">", sChr)
'  sName = Replace(sName, "|", sChr)
'  sName = Replace(sName, "&", sChr)
'  sName = Replace(sName, "%", sChr)
'  sName = Replace(sName, "*", sChr)
'  sName = Replace(sName, " ", sChr)
'  sName = Replace(sName, "{", sChr)
'  sName = Replace(sName, "[", sChr)
'  sName = Replace(sName, "]", sChr)
'  sName = Replace(sName, "}", sChr)
'  sName = Replace(sName, "!", sChr)
  sName = Replace(sName, "'", "")
End Sub

