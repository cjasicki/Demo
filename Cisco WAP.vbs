dim PingResult 
dim loopagain 
dim loopcount
dim strHost
dim ststage
dim IsAlive
dim TFTP
dim verdata2 
dim IsAlive1
dim Tstage
dim stage

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
            IsAlive1 = "NotFound"

         Case Else
            IsAlive1 = "Pass"

	End Select 

        fFile1.Close 

	
	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile1.ReadAll,TFTP)	
         Case 0
            TFTPERROR = "False"

         Case Else
            TFTPERROR = "True"

	End Select 
        fFile1.Close 
 

	
End Sub

 'Prompt for Salon Number
stitle = "POS Version 12202012 Setup Scrip"
sMsg = vbCrLf & "Please Enter the Salon number or Press Cancel to Quit." & vbCrLf & vbCrLf & vbCrLf & "Salon Number: " & generatedName
sMsg1 = "" & vbCrLf & vbCrLf & "Please Retype the Salon Number: " & generatedName
sMsg2 = "" & vbCrLf & vbCrLf & "Retype the Salon Number to comfirm: " & generatedName
tsti = a
Dim salonNum, salonNumTrim, tsti
    Do While True
        salonNum = InputBox( sMsg, sTitle, sDefault )
        salonNumTrim = Trim( salonNum )
          Select Case True
            Case IsEmpty( salonNum )
              Exit Do
            Case "" = salonNum
              sMsg = "Empty input not allowed" & vbCrLf & vbCrLf & sMsg1
              tsti = a
            Case "" = salonNumTrim
              sMsg = "Empty input not allowed" & vbCrLf & vbCrLf & sMsg1
              tsti = a
            Case 3 > Len( salonNumTrim )
              sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = a
	    Case 7 < Len( salonNumTrim )
              sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = a
	    Case IsNumeric (salonNum) = false
	      sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = a
            Case IsNumeric (salonNum) = true
	      sMsg = sMsg2
                if tsti = salonNum then
	          Exit Do
                end if
              tsti = salonNum
	    Case Else
              Exit Do
          End Select
    Loop

struser = "Cisco"
strpass = "Cisco"
loopcount = 0
Stage = 88
bx = "No"
verdata2 = "ap#"
ststage = 2
TFTP = "%Error opening tftp"
aces = "% Login invalid" 
txtdata = "#exit"

set WshShell = WScript.createObject ("WScript.Shell")
WshShell.run "C:\putty.exe -load cisco"
WScript.sleep 1000
WshShell.AppActivate strHost & " - PuTTY"
WScript.sleep 4000
WshShell.SendKeys "~"
WScript.sleep 5000
loopagain = true
Tstage = "Y"

do

WshShell.AppActivate strHost & " - PuTTY"

logfile()

Select Case Stage
 
  case 88
 
     if IsAlive1 = "NotFound" Then
	txtdata = "ap>"
	Stage = 1
     else
	msgbox "Make sure Device is Powered ON... and start over....."
        WshShell.SendKeys "%{F4}"
	WScript.sleep 100
        WshShell.SendKeys "{ENTER}"
	WScript.Quit
     end if


  case 1
     if IsAlive1 = "Pass" and Tstage = "Y" Then
    	WshShell.SendKeys "{ENTER}"
  	WScript.sleep 100   
        WshShell.SendKeys "en~"
        WScript.sleep 1500
	WshShell.SendKeys strpass & "~"
	loopagain = True
	txtdata = verdata2 
        Stage = ststage
        WScript.sleep 2000
    End if

     if Tstage = "N" Then
    	WshShell.SendKeys "{ENTER}"
        WScript.sleep 1500
	WshShell.SendKeys struser & "~"
        WScript.sleep 1500
	WshShell.SendKeys strpass & "~"
	loopagain = True
	txtdata = verdata2 
        Stage = ststage
        WScript.sleep 1500
     End if

     if IsAlive1 <> "Pass" then

        WScript.sleep 15000
        WshShell.SendKeys "{ENTER}" 
        if loopcount < 30 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           Stage = 99
        end if
      end if

  case 2
     if IsAlive1 = "Pass" Then     
         Stage = 3
         Loopagain = True
         Loopcount = 0
     else
         strpass = "abc123!!!"
         struser = "guest"
         verdata2 = "rta-" 
	 txtdata = "% Authentication failed"
         stage = 1
     end if


  case 3

        WScript.sleep 800
        WshShell.SendKeys "en~"
        WshShell.SendKeys "config t~"
        WshShell.SendKeys "ip ftp username ftpuser~"
        WshShell.SendKeys "ip ftp password Image!!!123~"
        WshShell.SendKeys "end~"
	WshShell.SendKeys "archive download-sw /overwrite ftp://10.1.97.62/ispc/c1140-k9w7-tar.124-21a.JY.tar~"  
	WScript.sleep 800
        WshShell.SendKeys "~"
        Stage = 4
        txtdata = "Configuring system to use new image...done"
  


  case 4

        if IsAlive1 = "Pass" Then  
           WScript.sleep 500
	   WshShell.SendKeys "copy ftp:wap_" & salonNum & ".txt startup-config"
           WshShell.SendKeys "{ENTER}"
	   WshShell.SendKeys "10.1.97.62~"
           WScript.sleep 5000
           WshShell.SendKeys "{ENTER}"
           WScript.sleep 5000
           WshShell.SendKeys "{ENTER}"
           WScript.sleep 5000
           WshShell.SendKeys "{ENTER}"
           txtdata = "bytes copied"
           Stage = 5
        else
           if loopcount < 60 then
               WScript.sleep 5000
	       loopagain = True
               loopcount = loopcount + 1  
 	   else
               WScript.Echo "TFTP Error.... (Stage 4)"
               Stage = 99
           end if
        End if

        if TFTPERROR = "True" Then  
            Stage = 3   
	    WScript.sleep 10000   
            TFTP = "nogoanddonothing123321"
        else
            TFTP = "%Error opening tftp"
        end if
   
  case 5

     if IsAlive1 = "Pass" Then  
        WScript.sleep 800
	WshShell.SendKeys "reload~"
        WScript.sleep 3000
        txtdata = "Proceed with reload?"
	Stage = 5.1
     else
        WScript.sleep 3000
        if loopcount < 30 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "FTP Failed. Ensure file exists. (Stage 5)"
           Stage = 99
        end if
     End if 

        if TFTPERROR = "True" Then  
            Stage = 4  
	    WScript.sleep 10000   
            TFTP = "nogoanddonothing123321"
        else
            TFTP = "%Error opening ftp"
        end if

    Case 5.1

     if IsAlive1 = "Pass" Then  
	WshShell.SendKeys "~"  
	WScript.sleep 800
        txtdata = "Username:"
        Stage = 6
        Loopagain = True
        Loopcount = 0

     else
      	WshShell.SendKeys "no~" 
	WScript.sleep 2000
	WshShell.SendKeys "~"
     End if   

  case 6

     WScript.sleep 15000
     if IsAlive1 = "Pass" Then 
         ststage = 7
         stage = 1

         Tstage = "N"
         Loopagain = True
         Loopcount = 0
     else
        WshShell.SendKeys "{ENTER}"
        if loopcount < 50 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "Reboot time-out error (Stage 6)"
           Stage = 99
        end if
     End if    

  case 7

     WScript.sleep 6000
     txtdata = "wap-00"
     logfile()

     if IsAlive1 = "Pass" Then  
	 Stage = 8
     end if

      if loopcount < 20 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "Failed to relogin after Reboot (Stage 6)"
           Stage = 99
        end if


  case 8

     if IsAlive1 = "Pass" Then  
	 msgbox "Check prompt for wap-" & salonNum & "-10. If Prompt is good... Access point is done" 

         WshShell.SendKeys "exit~"
         WScript.sleep 2000
         WshShell.SendKeys "%{F4}"
         WScript.sleep 1000
         WshShell.SendKeys "{ENTER}"
         WScript.Quit

        if loopcount < 20 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "Failed to setup router (Error Stage 8)"
           Stage = 99
        end if
     End if   


  case 99
        WScript.Echo "Error please start over. (Stage " & Stage & ")"
        WshShell.SendKeys "exit~"
        WScript.sleep 2000
        WshShell.SendKeys "%{F4}"
        WScrpt.sleep 1000
        WshShell.SendKeys "{ENTER}"
	WScript.Quit

  case else
	loopagain = False
     
End Select

Loop While loopagain

msgbox "If you see this something didn't go right s" & stage

        WshShell.SendKeys "exit~"
        WScript.sleep 2000
        WshShell.SendKeys "%{F4}"
        WScript.sleep 200
        Wshshell.Sendkeys "~"
         WScript.Quit