Public Sub SaveinSenderFolder()
Dim objOL As Outlook.Application
Dim objMsg As Outlook.MailItem 'Object
Dim objAttachments As Outlook.Attachments
Dim objSelection As Outlook.Selection
Dim i As Long
Dim lngCount As Long
Dim strFile As String
Dim strFolderpath As String, strFolder As String
Dim strDeletedFiles As String
         
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
 
'strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
                  strFolderpath = "file path"
   ' On Error Resume Next
 
Set objOL = Application
Set objSelection = objOL.ActiveExplorer.Selection
 
' The attachment folder needs to exist
' You can change this to another folder name of your choice
 
 
    ' Check each selected item for attachments.
    For Each objMsg In objSelection
 
    ' Set the Attachment folder.
   ' strFolder = strFolderpath & "\OLAttachments\"
     strFolder = strFolderpath
     
    Set objAttachments = objMsg.Attachments
    strFolder = strFolder & objMsg.SenderName & "\"
    
  ' if the sender's folder doesn't exist, create it
 If Not FSO.FolderExists(strFolder) Then
 FSO.CreateFolder (strFolder)
 End If
 
    lngCount = objAttachments.Count
         
    If lngCount > 0 Then
     
    ' Use a count down loop for removing items
    ' from a collection. Otherwise, the loop counter gets
    ' confused and only every other item is removed.
     
    For i = lngCount To 1 Step -1
     
    ' Get the file name.
    strFile = objAttachments.Item(i).FileName
    
    ' This code looks at the last 4 characters in a filename
      sFileType = LCase$(Right$(strFile, 4))
      
    Select Case sFileType
 ' Add additional file types below
       Case ".jpg", ".xml", ".gif"
        If objAttachments.Item(i).Size < 5200 Then
     GoTo nexti
        End If
      End Select
     
    ' Combine with the path to the folder.
    strFile = strFolder & strFile
     
    ' Save the attachment as a file.
    objAttachments.Item(i).SaveAsFile strFile
        
nexti:
    Next i


    End If
     
    Next
     
ExitSub:
   
Set objAttachments = Nothing
Set objMsg = Nothing
Set objSelection = Nothing
Set objOL = Nothing

End Sub


fice.net‚*.roaming.osi.office.net‚*.osi.office.net0UÅ0åÀ ~Û
Ïèã ŒKèÖ\m0U#0€XˆŸÖÜœH"·>ÿ„ˆèæ…ÿú}0¬U¤0¡0ž › ˜†Khttp://mscrl.microsoft.com/pki/mscorp/crl/Microsoft%20IT%20TLS%20CA%201.crl†Ihttp://crl.microsoft.com/pki/mscorp/crl/Microsoft%20IT%20TLS%20CA%201.crl0…+y0w0Q+0†Ehttp://www.microsoft.com/pki/mscorp/Microsoft%20IT%20TLS%20CA%201.crt0"+0†http://ocsp.msocsp.com0>	+‚710/'+‚7‡Ú†uƒîÙ‚É…µža…ôë`]„ÒßB‚ç“zd0MU F0D0B	+‚7*0503+'http://www.microsoft.com/pki/mscorp/cps0'	+‚7
00
+0
+0
	*†H†÷
‚7ÏÎ·#²¯È©$0€+à–^Ýø„{—²jkýôÊwªk;_”TËDGlë«Ÿ¦&§àqCŸØ«òÔ}±ÅsDÊþ4Ý]óV‰ieg°|Ùyµ¸Lf|š¤rKs¼©ù._®@q…”|éÚµm²Ou?A;QÈ†÷” Þ’J^6aÝüƒøz\š¬^\b£3IÓƒ>·<¿†‡8«KÎ¡ÆÖß4Kž9ÕÀZÏæõMu7ïotg(½‡ ³Ž<è®	‡‰÷5\is’7YrÔ<×˜'§_‹¤‹jq‹½›=’º`5Ž>	îô~#²ñÞW²¢Ê+Â2·¥²ïÇV§97ß
Eðà'$|ø'U–™ÂÒqG‚>§–/§âZÍÞ#ÝbvÜwC4ÝÀ–8ÇÿMË÷ð$Ï¤¶<?ì´½­~3ýSIž„Ï¿àË ŸZ_Sh=ð8:ª<uItãð‹§c`ö…ÿ—ï‹ÿŠõœÎ»\‡wï«HÓýéÌømHlä»Z6X:4ØÏ¤Æ-zC‘Ê^%)lú¥)‰-òûµA¦µÄød5$™—Ônýº$xs ÞjýD3yµ ìELÿ¸Naj@ØŠøÍ›Å”Ij+Ž±¤R—c? ÔuñYÍkÛîü¤«
AW±K8üvÝAd¸0‚´0‚œ ¸zP¾œÚ-M>9Q¿U0
	*†H†÷
0Z10	UIE10U
	Baltimore10U
CyberTrust1"0 UBaltimore CyberTrust Root0
160520125128Z
240520125128Z0‹10	UUS10U
Washington10URedmond10U
Microsoft Corporation10UMicrosoft IT10UMicrosoft IT TLS CA 10‚"0
	*†H†÷
‚0‚
‚Žóñ„uw¼¾É¤õ¥S+P`™ÜØ}Ù$µárI7HþÚ†“£}šKMwy~e©|n6âGÔ6IÐÌ$'â®qÍÙWt<V’J‡bùã^ÞÑA¦@¢,ŽU×&ØhB«ì
ÞÝ^a•³¬lÎ¨à­¯_Ê¦äQh.'ýT*q¤Ì»~’ñöSQ1Ð‚°ÊcÐòJÍôiKôZ–V9&ÉK
c42€¥åê(·ÁÀmð(ÔJ€¬sØõ/­®—b,þç¥ð­AS+1Ãs!ÝãxcœQ†µ¢HIÀyC+™¸KàÇWlÄ¹¦T-Sò£ò¤ÌÙÁTóˆÍäLÓŠ3!X­¿¾9p9i7ø[¤cï	ß‡bÇšÅÑ^bvÝiæ8»D°-M:Æ.€à`r[òÌæŽ\:!]°9+½êË–Ë¶jtŒŽ¼Ÿ¥@A‘\v§ñª:‡·wßqj/(aB-r«gàÁ{Ëœ@šÄ
DŒ|=º©±Ù"½_AjsTöfâ ù¦9ÞÒwói¿·»<Ä“ÿrµ6hßú9O²ö¹ÿ­íã†|òHIw—öýöO¦^oç9ÀQe=ë±—0ìêžsü, ß‹Ïs_µB&•ìØ©'ÐÐe”\â?Ê()¥ ôÈÝ]´ÿæÝ£‚B0‚>0UXˆŸÖÜœH"·>ÿ„ˆèæ…ÿú}0U#0€åY0‚GXÌ¬úT6†{:µMð0Uÿ0ÿ0Uÿ†0'U% 0+++	04+(0&0$+0†http://ocsp.digicert.com0:U3010/ - +†)http://crl3.digicert.com/Omniroot2025.crl0=U 60402U 0*0(+https://www.digicert.com/CPS0
	*†H†÷
‚0šÆjýï“¾‚wùv mž{0#{¨)Zôj>Ç–ß¸KRä
œ8íxcµsÀ;à§ÿIQ•2¸Ð›©åÏ–€ÕJaþÄjÆßAF"œ€fëB äó¤!£˜ÐztöŒèÃÒ+ª+ÎYDç\	Bë×ýM¹olD5&‡º£;h°ç ÉóÌ«Ÿ•PË®d€»‡
]Î¦k²}ã=6â)Q·%üÐ	ã°­Äb.>~…&²ö¯÷m1sÆ˜©r“ÎÊ=<ìÙpè€õ«xj‡MÆ7¨
v¨ï`|p<8×3Lä7eû‘³èva*eõX”³EïÀO{¸ø¯ãwG%ð;Œ€i%LMEMˆðó‹i/˜@Ž±0Øf
