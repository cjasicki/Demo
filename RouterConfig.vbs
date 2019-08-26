on error resume next

'connects to network folder
'net use R: \\SPALMS02\Config-Archive /user svc_posups shoppc3
Set objNetwork = WScript.CreateObject("WScript.Network")
objNetwork.MapNetworkDrive "R:", "\\SPALMS02\Config-Archive", False, "user", "pass"

'Figures out the file name for the router and AP config files
salon = InputBox("Enter Salon Number:") 
salonLen = (Len(salon))

Select Case salonLen
  Case 3
	strFileCopy1 = "rta-00000" & salon & "-01.regisbb.com-Running.Config"
	strFileCopy2 = "wap-00000" & salon & "-01.regisbb.com-Running.Config"
        strFileCopy3 = "swa-00000" & salon & "-01.regisbb.com-Running.Config"
  Case 4
	strFileCopy1 = "rta-0000" & salon & "-01.regisbb.com-Running.Config"
	strFileCopy2 = "wap-0000" & salon & "-01.regisbb.com-Running.Config"
        strFileCopy3 = "swa-0000" & salon & "-01.regisbb.com-Running.Config"
  Case 5
	strFileCopy1 = "rta-000" & salon & "-01.regisbb.com-Running.Config"
	strFileCopy2 = "wap-000" & salon & "-01.regisbb.com-Running.Config"
        strFileCopy3 = "swa-000" & salon & "-01.regisbb.com-Running.Config"
  Case 6
	strFileCopy1 = "rta-00" & salon & "-01.regisbb.com-Running.Config"
	strFileCopy2 = "wap-00" & salon & "-01.regisbb.com-Running.Config"
        strFileCopy3 = "swa-00" & salon & "-01.regisbb.com-Running.Config"
  Case else
        msgbox "    Bad Salon#" & Chr(10) & "Please try again."
        Wscript.Quit

End Select

	strFilePast1 = "router_" & salon & ".txt"
	strFilePast2 = "wap_" & salon & ".txt"
	strFilePast3 = "switch_" & salon & ".txt"

strFileEX = 0
strFileCo = 0 
strcount = 0
stpass = 0
strcopy = strFileCopy1
strpast = strFilePast1
strnextlast = 10000000

Do While stpass = 0
'looking for last created folder
Set fs = CreateObject("Scripting.FileSystemObject")
Set MainFolder = fs.GetFolder("\\SPALMS02\Config-Archive\")
For Each fldr In MainFolder.SubFolders
    If fldr.DateLastModified > LastDate and fldr.DateLastModified < strnextlast Or IsEmpty(LastDate) Then
        LastFolder = fldr.Name
        LastDate = fldr.DateLastModified
    End If
Next

   If fs.FileExists ("\\SPALMS02\Config-Archive\" & LastFolder & "\" & strcopy) Then 
      fs.CopyFile "\\SPALMS02\Config-Archive\" & LastFolder & "\" & strcopy, "\\spatftp01\TFTP-Root\" & strpast , false
      strcount = 0
      strnextlast = 10000000 
      LastDate = Empty
      

      If Err.Number = 0 Then
         strupste = "copied " & strcopy & " from " & LastFolder & " Folder on: " & now()
         strFileCo = strFileCo + 1
      Else
         strupste = err.description & " " & strcopy & " on destination folder " & now()
         strFileEX = strFileEX + 1
      End If 
      
      Err.Clear
      Set ObjFSO = CreateObject("Scripting.FileSystemObject")
      Set objLog = objFSO.OpenTextFile("K:\POS Database\CiscoCopyLog\CiscoCopylog.txt", 8, True, 0)
      objLog.WriteLine strupste
      objLog.Close

      If strcopy = strFileCopy1 then
         strcopy = strFileCopy2
         strpast = strFilePast2 
      Else
	 strcopy = strFileCopy3
         strpast = strFilePast3
      End if

   Else
      strnextlast = LastDate 
      LastDate = Empty
      strcount = strcount + 1

      IF strcount = 20 then
        strnextlast = 10000000
        LastDate = Empty
	strcount = 0
   	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	Set objLog = objFSO.OpenTextFile("K:\POS Database\CiscoCopyLog\CiscoCopylog.txt", 8, True, 0)
	objLog.WriteLine "Could not Find: " & strcopy & chr(10) &". Last folder checked was: " & LastFolder & " Date: " & now()
	objLog.Close
        
	  if strcopy = strFileCopy1 then
             strcopy = strFileCopy2
             strpast = strFilePast2
          ElseIf strcopy = strFileCopy2 then
	     strcopy = strFileCopy3   
             strpast = strFilePast3 
          else
             stpass = 1
 	  end if
      end if
   end if
Loop
msgbox "Files copied: " & strFileCo & Chr(10) & "Files not copied because file already exists: " & strFileEX 
'objNetwork.RemoveNetworkDrive "R:", True, True
