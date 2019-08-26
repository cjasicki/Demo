dim PingResult 
dim loopagain 
dim loopcount
dim IsAlive
dim IsAlive1 
dim Erro
dim Stage
dim txtdata
dim strHost
dim sTempFile1
dim sleeptimer
dim alreadyupdated
dim bx

Function testping()
     Const OpenAsASCII = 0 
     Const FailIfNotExist = 0 
     Const ForReading =  1 
     Dim objShell, objFSO, sTempFile, fFile 

	Set objShell = CreateObject("WScript.Shell") 
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
	sTempFile = objFSO.GetSpecialFolder(2).ShortPath & "\" & objFSO.GetTempName 
	objShell.Run "%comspec% /c ping.exe -n 1 -w 500 " & strHost & ">" & sTempFile, 0 , True 
	Set fFile = objFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile.ReadAll, "TTL=") 
         Case 0
            IsAlive = False 

         Case Else
            IsAlive = True

	End Select 
	fFile.Close 
	objFSO.DeleteFile(sTempFile)

    Set objFSO = Nothing
    Set objShell = Nothing

End Function

Sub logfile()

  Dim objShell1, objFSO, sTempFile1, fFile1
     Const OpenAsASCII = 0
     Const FailIfNotExist = 0 
     Const ForReading =  1 
	Set objShell1 = CreateObject("WScript.Shell") 
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
        sTempFile1 = "c:\putty.log" 
	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile1.ReadAll,txtdata) 
         Case 0
            IsAlive1 = "False"

         Case Else
            IsAlive1 = "True"

	End Select 

        fFile1.Close 

	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile1.ReadAll,"-sh: /var/root/update.sh: not found")	
         Case 0
            Erro = "False"

         Case Else
            Erro = "True"

	End Select 
        fFile1.Close 
	
	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile1.ReadAll,"Already updated")	
         Case 0
            alreadyupdated = "False"

         Case Else
            alreadyupdated = "True"

	End Select 
        fFile1.Close 
	
	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile1.ReadAll,"Ramdisk2 alreay updated")	
         Case 0
            alreadyupdated2 = "False"

         Case Else
            alreadyupdated2 = "True"

	End Select 
        fFile1.Close 
	
    Set objFSO = Nothing
    Set objShell1 = Nothing

End Sub
Sub timer()
   WScript.sleep sleeptimer
End Sub

strHost = "192.168.86.1"
loopagain = False
loopcount = 0
Stage = 1
txtdata = "na"
bx = "No"

set WshShell = WScript.createObject ("WScript.Shell")
WshShell.run "C:\putty.exe " & strHost & " -l root -pw Snprysm2"
WScript.sleep 2000
WshShell.AppActivate strHost & " - PuTTY"
WScript.sleep 1000
WshShell.SendKeys "N%~" 
WScript.sleep 3000
WshShell.SendKeys "~" 
WScript.sleep 3000
WshShell.SendKeys "iptables -F~"
WshShell.SendKeys "iptables -F -t nat~"
WshShell.SendKeys "iptables -F -t mangle~"
WshShell.SendKeys "cd /var~"
WshShell.SendKeys "tftp -g -r PrysmPro3.3.0.0.gz 192.168.86.2~"
WshShell.SendKeys "tftp -g -r PrysmPro3.3.0.0.gz.md5 192.168.86.2~" 
WshShell.SendKeys "md5sum -c PrysmPro3.3.0.0.gz.md5~"
sleeptimer = 5000

Do
   loopcount = loopcount + 1
   timer() 
   logfile()
   testping()

     if Erro = "True" then
    	WScript.Echo "-sh: /var/root/update.sh: not found"
 	Stage = 999
     end if

     if alreadyupdated = "True" and bx = "No" then
    	'WScript.Echo "u-boot Already updated"
	sleeptimer = 5000
 	Stage = 9
     end if

     if alreadyupdated2 = "True" Then
    	'WScript.Echo "Already updated"
	sleeptimer = 5000
 	Stage = 7
        txtdata = " "
     end if

Select Case Stage
  case 1

     if IsAlive = True Then
        Stage = 2
	txtdata = "PrysmPro3.3.0.0.gz: OK"
	IsAlive1 = "na"
	loopagain = True
     Else
        if loopcount < 15 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 1)"
           Stage = 999
        end if
     End if

  case 2
     if IsAlive = True and IsAlive1 = "True" Then
	WScript.sleep 2000
        WshShell.SendKeys "gunzip PrysmPro3.3.0.0.gz~"
	WshShell.SendKeys "tar  -xvf   /var/PrysmPro3.3.0.0~"
	WshShell.SendKeys "chmod  777  /var/root/update.sh~"
	WshShell.SendKeys "/var/root/update.sh~"

        loopcount = 2
        Stage = 3
	txtdata = "uImage3.3.0.0: OK"
	IsAlive1 = "na"
	loopagain = True
	sleeptimer = 30000
     Else
        if loopcount < 16 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 2)"
           Stage = 999
        end if
     End if

  case 3
    if IsAlive = True and IsAlive1 = "True" Then
        loopcount = 2
        Stage = 4
	txtdata = "Ramdisk3.3.0.1: OK"
	IsAlive1 = "na"
	loopagain = True
      else
        if loopcount < 22 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 3)"
           Stage = 999
        end if
     end if

  case 4
      if IsAlive = True and IsAlive1 = "True" Then
        loopcount = 2
        Stage = 5
	txtdata = "uboot3.3.0.0: OK"
	IsAlive1 = "na"
	loopagain = True
      else
        if loopcount < 22 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 4)"
           Stage = 999
        end if
     end if

  case 5
     if IsAlive = True and IsAlive1 = "True" Then
        loopcount = 2
 	txtdata = "ÿ#"
	IsAlive1 = "na"
	loopagain = True
	sleeptimer = 3000
 	Stage = 10
      else
        if loopcount < 36 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 5)"
           Stage = 999
        end if
     end if

  case 6
     if IsAlive = True and IsAlive1 = "True" Then

        WshShell.SendKeys "~"
	WScript.sleep 15000
        WshShell.SendKeys "~"
	WScript.sleep 5000
        WshShell.SendKeys "%{F4}"
        WScript.sleep 600
        WshShell.SendKeys "Y%"
        WScript.sleep 800
        set WshShell = WScript.createObject ("WScript.Shell")
	WshShell.run "C:\putty.exe " & strHost & " -l root -pw Snprysm2"
	WScript.sleep 1000
	WshShell.SendKeys "N%" 
        WScript.sleep 1000
	WshShell.AppActivate strHost & " - PuTTY"
        WScript.sleep 1000
	WshShell.SendKeys "~" 
	WScript.sleep 1000
	WshShell.SendKeys "cd /var~"
	WshShell.SendKeys "tftp -g -r RamdiskU3.3.0.1 192.168.86.2~"
	WshShell.SendKeys "tftp -g -r RamdiskU3.3.0.1.md5  192.168.86.2~"
	WshShell.SendKeys "md5sum -c RamdiskU3.3.0.1.md5~"
	WshShell.SendKeys "update_ramdisk2 RamdiskU3.3.0.1~"
	
	loopcount = 2
	Stage = 7 
        txtdata = "update_version  done"
	IsAlive1 = "Now using ALT RAMSISK"
	loopagain = True
  	sleeptimer = 30000

     else
        if loopcount < 60 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 6)"
           Stage = 999
        end if
     end if

  case 7
     if IsAlive = True and IsAlive1 = "True" Then
    
 	WScript.sleep 30000     
        WshShell.SendKeys "cls~"
        WshShell.SendKeys "get_version~"        
        WScript.Echo "Check that version listed is:" & chr(13) & chr(13) & "u-boot = 3.3.0.0" & chr(13) & "RAMDISK = 3.3.0.1" & chr(13) & "RAMDISK UPDATE = U3.3.0.1" & chr(13) & "LINUX = 3.3.0.0" & chr(13) & "LINUX UPDATE =" 
 	WScript.sleep 1200
	WshShell.SendKeys "exit~"
        WScript.sleep 600
       	Stage = 99
      else
        if loopcount < 42 then
	   loopagain = True 
 	else
           WScript.Echo "Lost Connect to Prysm Pro. Please start over."
           Stage = 999
        end if
     end if

  case 9
     if IsAlive = False Then
      
        loopcount = 2
        Stage = 6
	txtdata = " "
	IsAlive1 = "na"
	loopagain = True
	bx = "Yes"
        WshShell.SendKeys "~"
	sleeptimer = 30000

      else
        if loopcount < 200 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 9)"
           Stage = 999
        end if
     end if

  case 10
     if IsAlive = False and IsAlive1 = "True" Then
      
        loopcount = 2
        Stage = 6
	txtdata = " "
	IsAlive1 = "na"
	loopagain = True
	bx = "Yes"
        WshShell.SendKeys "~"
	sleeptimer = 30000
     
      else
        if loopcount < 75 then
	   loopagain = True 
 	else
           WScript.Echo "Can't Connect to Prysm Pro. Please start over. (Stage 10)"
           Stage = 999
        end if
     end if

  case 99
	WScript.Quit

  case else
	loopagain = False
     
End Select

Loop While loopagain

msgbox "If you see this something didn't go right"