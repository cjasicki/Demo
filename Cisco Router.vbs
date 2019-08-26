'verion 7.12 Chad Jasicki

'Change log
'7.5 removed LTE firmware update. 
'7.7 misc fixes
'7.6 Changed ISO to "c1900-universalk9-mz.SPA.153-3.M2.bin"
'7.8 changed to work with updated ver 15.0 config files
'7.9 updated IOS to c1900-universalk9-mz.SPA.153-3.M4.bin
'7.10 update for Bell Wic card - need to fix WICFRIMWARE verable and firmware copy 
'7.11 update to copy CHECK_PACKET_LOSS.tcl to root
'7.12 updated IOS to c1900-universalk9-mz.SPA.155-3.M2.bin
'7.13 EHWIC lte version check rem'ed out
'7.14 EHWIC LTE version check reworked and enabled 7/28/2016
'7.15 fixed reload issue, where it wouldn't log in after reboot
'7.16 changed when format command was issued, now it is after LTE WIC card version check

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
dim passtry
dim filetran
dim IOSVER
dim AFRIMWARE 
dim VFRIMWARE 
dim tx123
dim Newtx
dim Newtx1
dim Newtx2
dim Newtx3
dim TFTPERROR
dim TFTPERROR1 
dim strcpyfile 
dim intResult
dim WIC 
dim UNTformat 
dim strrunWicFlash
dim miccount
dim strfim
dim EHWICFIRMWARE
dim strcyc

strcyc = 0
UNTformat = 1
miccount = 1
Newtx = 1
Newtx1 = 1
Newtx2 = 1
Newtx3 = 1
tx123 = 0
tx1234 = 0 
tx1235 = 0
strrunWicFlash = 0
tx1236 = 0
strfim = 0
EHWICFIRMWARE = 0

'checks ping results and based on that assume the router is ready or not, the start of the script and suring the reboot
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

strfound = 0
strloopagain = True

  Dim objShell1, objFSO, sTempFile1, fFile1
     Const OpenAsASCII = 0
     Const FailIfNotExist = 0 
     Const ForReading =  1 
	Set objShell1 = CreateObject("WScript.Shell") 
	Set objFSO = CreateObject("Scripting.FileSystemObject") 
        sTempFile1 = "c:\putty.log" 
	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	strCharacters = fFile1.ReadAll
	strCharacters = Replace(strCharacters, vbCrLf, "")
  Do
        tx123 = InStr(Newtx,strCharacters,txtdata,1)

	if tx123 = 0 Then
 	  IsAlive1 = "Not Found"
	  strloopagain = False
 	Else
	  IsAlive1 = "Pass"
          Newtx = tx123 + 1
	  strloopagain = True
	  strfound  = 1
	end if 

  Loop While strloopagain
        fFile1.Close 

	if strfound  = 1 Then
 	  IsAlive1 = "Pass"
	end if
	strfound = 0
	strloopagain = True

	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	strCharacters = fFile1.ReadAll
	strCharacters = Replace(strCharacters, vbCrLf, "")
  Do
	tx1234 = InStr(Newtx1,strCharacters,TFTP,1)	
      	if tx1234 = 0 Then
            TFTPERROR = "False"
	    strloopagain = False
	Else
            TFTPERROR = "True"
	    Newtx1 = tx1234 + 1
	    strloopagain = True
	    strfound  = 1

	End if

  Loop While strloopagain
        fFile1.Close 

	if strfound  = 1 Then
 	  TFTPERROR = "True"
	end if


	strfound = 0
	strloopagain = True

	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	strCharacters = fFile1.ReadAll
	strCharacters = Replace(strCharacters, vbCrLf, "")
  Do
	tx1235 = InStr(Newtx2,strCharacters,TFTP1,1)	
	if tx1235 = 0 Then
            TFTPERROR1 = "False"
	    strloopagain = False
	Else
            TFTPERROR1 = "True"
	    Newtx2 = tx1235 + 1
	    strloopagain = True
	    strfound  = 1

	End if 

  Loop While strloopagain
        fFile1.Close 

	if strfound  = 1 Then
	TFTPERROR1 = "True"
	end if

	strfound = 0
	strloopagain = True

	Set fFile1 = objFSO.OpenTextFile(sTempFile1, ForReading, FailIfNotExist, OpenAsASCII) 
	strCharacters = fFile1.ReadAll
	strCharacters = Replace(strCharacters, vbCrLf, "")
  Do
	tx1236 = InStr(Newtx3,strCharacters,TFTP2,1)	
	if tx1236 = 0 Then
            TFTPERROR3 = "False"
	    strloopagain = False
	Else
            TFTPERROR3 = "True"
	    Newtx3 = tx1236 + 1
	    strloopagain = True
	    strfound  = 1

	End if 

  Loop While strloopagain
        fFile1.Close 

	if strfound  = 1 Then
	   TFTPERROR3 = "True"
	end if


End Sub

'Prompt for Salon Number
stitle = "POS Version 7.7 Setup Scrip"
sMsg = vbCrLf & "Please Enter the Salon number or Press Cancel to Quit." & vbCrLf & vbCrLf & vbCrLf & "Salon Number: " & generatedName
sMsg1 = "" & vbCrLf & vbCrLf & "Please Retype the Salon Number: " & generatedName
sMsg2 = "" & vbCrLf & vbCrLf & "Retype the Salon Number to comfirm: " & generatedName
tsti = a
Dim salonNum, salonNumTrim, tsti
    Do While True
        salonNum = InputBox(sMsg, sTitle, sDefault)
        salonNumTrim = Trim(salonNum)
          Select Case True
            Case IsEmpty(salonNum)
              WScript.Quit
	      'Exit Do
            Case "" = salonNum
              sMsg = "Empty input not allowed" & vbCrLf & vbCrLf & sMsg1
              tsti = a
            Case "" = salonNumTrim
              sMsg = "Empty input not allowed" & vbCrLf & vbCrLf & sMsg1
              tsti = b
            Case 3 > Len( salonNumTrim )
              sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = c
	    Case 7 < Len( salonNumTrim )
              sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = d
            Case "TEST" = UCase(salonNum)
	      salonNum = "Test" 
              Exit Do
	    Case IsNumeric (salonNum) = false
	      sMsg = "Invalid Salon Number" & vbCrLf & vbCrLf & sMsg1
              tsti = e
            Case IsNumeric (salonNum) = true
	      sMsg = sMsg2
                if tsti = salonNum then
	          Exit Do
                end if
              tsti = salonNum
	    Case Else
	
          End Select
    Loop

passtry = 0

IOSVER = "c1900-universalk9-mz.SPA.155-3.M2.bin"
AFRIMWARE = "MC7700_ATT_03.05.10.02_00.cwe"
VFRIMWARE = "verizon_wicfirmware.cwe"

struser = "cisco"
strpass = "cisco"
loopcount = 0
passtry = 0
Stage = 88
verdata2 = "yourname#"
ststage = 2
TFTP = "%Error opening ftp"
TFTP1 = "%Warning:There is a file already"
TFTP2 = "No such file or directory"
aces = "% Login invalid" 
txtdata = "exit"
loopagain = true
set WshShell = WScript.createObject ("WScript.Shell")
WshShell.run "C:\putty.exe -load cisco"
WScript.sleep 10000
WshShell.SendKeys "~"
WScript.sleep 3000

do
WshShell.AppActivate ("COM1 - PuTTY")
logfile()
Select Case Stage
  case 88
     if IsAlive1 <> "Pass" Then
	txtdata = "name:"
	Stage = 1
     else
	msgbox "Make sure Device is Powered ON... and try again."
	Stage = 99
     end if

  case 1
     if IsAlive1 = "Pass" Then
	WScript.sleep 200  
  	WshShell.SendKeys struser & "~"
	WScript.sleep 200
	WshShell.SendKeys strpass & "~"
	loopagain = True
	txtdata = verdata2 
        Stage = 2
        WScript.sleep 2000

     Else
        WScript.sleep 4000
        WshShell.SendKeys "{ENTER}" 
        if loopcount < 90 then
	   loopagain = True
           loopcount = loopcount + 1  

 	else
           Stage = 99 
           msgtxt = "login Error# Stage 1"
        end if
     End if

  case 2
     if IsAlive1 = "Pass" Then    
         Stage = 3
         loopagain = True
         loopcount = 0
         passtry = 5
     else
	 txtdata = "name:"
	 passtry = passtry + 1 
         verdata2 = "rta-"
     end if

Select Case passtry
    Case 1
           strpass = "pass1"
           struser = "user1"
           WScript.sleep 2000
           Stage = 1
    Case 2
           strpass = "pass2"
           struser = "user2"
           WScript.sleep 2000
           Stage = 1
    Case 3
           strpass = "pass3"
           struser = "user3"
           WScript.sleep 2000
           Stage = 1
    Case 6
           msgtxt = "login Error# User or Pass not Valid"
           Stage = 99

    Case Else
       
    End Select

  case 3
	WScript.sleep 1000
        WshShell.SendKeys "en~"
	WScript.sleep 500
	WshShell.SendKeys "gooseberry123!!!~"
	WScript.sleep 500
    if strcyc = 1 then
	Stage = 7
	UNTformat = 100
    else
        Stage = 4
    end if

  case 4
	WshShell.SendKeys "config t~"
	WshShell.SendKeys "int gig0/0~"
        WshShell.SendKeys "shut~"
	WshShell.SendKeys "int gig0/1~"
	WshShell.SendKeys "ip address dhcp~"
	WScript.sleep 7000
	WshShell.SendKeys "no shut~"
	WshShell.SendKeys "ip ftp username UserID~"
	WshShell.SendKeys "ip ftp password Pass~"
	WshShell.SendKeys "ip route 0.0.0.0 0.0.0.0 dhcp~"
	WshShell.SendKeys "ip ftp source-interface gi0/1~"
	WshShell.SendKeys "end~"
	WScript.sleep 2500
	WshShell.SendKeys "sh inv~"
	WScript.sleep 2500
	txtdata = "WAN Interface Card - HWIC CSU/DSU"
        Stage = 31
	loopcount = 0
        FTPTR = 1    

   case 31
     WScript.sleep 500
     if IsAlive1 = "Pass" Then 
	WIC = 0
        Stage = 35
   else
        txtdata = "EHWIC-4G-LTE-AT"
        Stage = 32
	EHWICFIRMWARE = 1
     end if        

    case 32
Select Case EHWICFIRMWARE
   case 1
     WScript.sleep 500
     if IsAlive1 = "Pass" Then 
         strfim = "05.05.58.00" 
'        flashstr = "microcode reload cellular 0 0 modem-provision flash:MC7700_ATT_03.05.10.02_00.cwe~"
'        WICFRIMWARE = AFRIMWARE
	WIC = 1
        Stage = 33
     else
        txtdata = "EHWIC-4G-LTE-VZ"
        EHWICFIRMWARE = 2
     end if
   
   case 2
       WScript.sleep 500
       if IsAlive1 = "Pass" Then 
          strfim = "05.05.58.01"
	  WIC = 1
          Stage = 33
       else
        txtdata = "EHWIC-4G-LTE-V"
        EHWICFIRMWARE = 3
       end if 

   case 3
       WScript.sleep 500
       if IsAlive1 = "Pass" Then 
          strfim = "03.05.10.06"
	  WIC = 1
          Stage = 33
       else
        txtdata = "EHWIC-4G-LTE-A"
        EHWICFIRMWARE = 4
       end if 

   case 4
       WScript.sleep 500
       if IsAlive1 = "Pass" Then 
          strfim = "03.05.29.02"
	  WIC = 1
          Stage = 33
'          flashstr = "microcode reload cellular 0 0 modem-provision flash:verizon_wicfirmware.cwe~"
'          WICFRIMWARE = VFRIMWARE
       else
	  msgtxt  = "Error... Couldn't detect LTE Wic Card, check if LTE WIC card is powered and IOS version is correct" & vbCrLf & vbCrLf
          Stage = 99
       end if 

   case else

End Select

   case 33 
        if WIC = 1 Then
	      WshShell.SendKeys "sh cellular 0/0/0 hardware~"	
	      WScript.sleep 4500
	      txtdata = strfim 
	      WScript.sleep 4500
              Stage = 34
              'FTPTR = 33 not working with new Cell card	 
         else
              Stage = 35   
          end if

     case 34
          if IsAlive1 = "Not Found" Then 
              msgtxt  = "LTE WIC Card dose not have the correct firmware version, FTPTR case# 34....." & vbCrLf & vbCrLf 
	      N0Copy = 1
              Stage = 99  'was 3434

	      'filetran = "copy ftp flash:"
	      'strcpyfile = "ISPC/" & WICFRIMWARE
              'strrunWicFlash = 1
	      'N0Copy = 0
              'loopcount = 0
	  else
              Stage = 35
	      N0Copy = 1
              FTPTR = 1
          end if
 	        
	case 3434
          N0Copy = 1
          if WIC = 1 & strrunWicFlash = 1 Then
	      WshShell.SendKeys flashstr 
	      WScript.sleep 800
	      WshShell.SendKeys "~"
	      WScript.sleep 800
	      WshShell.SendKeys "~"
	      WScript.sleep 5000
              txtdata = "Modem in HWIC slot 0/0 is now up" 
              Stage = 322
          end if 

   case 35
Select Case UNTformat
 case 1
     if salonNum = "Test" Then
	stage = 36
     else
	WshShell.SendKeys "format flash0:~"
	WScript.sleep 1000
	WshShell.SendKeys "~"
	WScript.sleep 1000
	WshShell.SendKeys "~"
	txtdata = "Format of flash0: complete"
        UNTformat = 2
      end if
 
 case 2
     if IsAlive1 = "Pass" Then
	   WshShell.SendKeys "mkdir flash0:archive~"
	   WScript.sleep 500
	   WshShell.SendKeys "~"
	   WScript.sleep 500
	   WshShell.SendKeys "~"
	   WScript.sleep 500
	   WshShell.SendKeys "~"
           stage = 36
           UNTformat = 9
     
        if loopcount < 90 then
	   loopagain = True
           loopcount = loopcount + 1  
     	else
           Stage = 99 
           msgtxt = "Format Error"
        end if
     End if

  case else
      'do nothing

End Select

'Copy cammands are issued here, curretly 3 of them - IOS, startup-config, and check_packet_loss.tcl
   case 36
     N0Copy = 1
     Select Case FTPTR
        case 1 
              If salonNum <> "Test" then
                  filetran = "copy ftp flash:"
                  strcpyfile = "ISPC/" & IOSVER
	          N0Copy = 0
                  loopcount = 0
              else
                  FTPTR = 2
              end if

        case 2
                 if salonNum = "Test" then
                 filetran = "copy ftp startup-config"
                 strcpyfile = "ISPC/Router_" & salonNum & ".txt"
	         N0Copy = 0
                 loopcount = 0
              else
                 filetran = "copy ftp startup-config"
                 strcpyfile = "router_" & salonNum & ".txt"
	         N0Copy = 0
                 loopcount = 0
 	      end if

        case 3
             If salonNum <> "Test" then
                  filetran = "copy ftp flash:"
                  strcpyfile = "ISPC/CHECK_PACKET_LOSS.tcl"
	          N0Copy = 0
                  loopcount = 0
             else
             	  Stage = 444
         	  N0Copy = 1
	     end if
	     
        case 4
	      Stage = 444
              N0Copy = 1
	      
        case else

           msgtxt  = "Error with an FTPCopy....." & vbCrLf & vbCrLf 
           Stage = 99

     End Select

        If N0Copy = 0 Then
     	     WshShell.SendKeys filetran
             WshShell.SendKeys "{ENTER}"
	     WshShell.SendKeys "10.1.97.62~"
	     WshShell.SendKeys strcpyfile
	     WshShell.SendKeys "~"
             txtdata = "bytes copied in"
             Stage = 333
	     loopcount = 0
             N0Copy = 1
         End If
    
    case 322

        WScript.sleep 6500

 	 'this checks to see if the microcode reload cellular firmware update was complete
          if IsAlive1 = "Pass" Then 
             miccount = miccount +1
             WScript.sleep 3000
             Stage = 3434
             FTPTR = 8
          else
               if loopcount < 90 then
	          loopagain = True
                  loopcount = loopcount + 1
                else
                  WScript.Echo "Microcde firmware failed (Stage 322a)"
		  Stage = 99
                end if
          End if                    
	  
    case 333 
 	 'this checks to see if file already exists and cancels the copy       
         if TFTPERROR1 = "True" Then  
	   WScript.sleep 500   
    	   WshShell.SendKeys "n"
           Stage = 36
           FTPTR = FTPTR + 1  
	   loopagain = True
         end if

          'Checks to see if copy was done Successfully.
          if IsAlive1 = "Pass" Then 
             TFTP = "%Error opening ftp"
             Stage = 36
             FTPTR = FTPTR + 1
          else
               if loopcount < 90 then
	          loopagain = True
                  loopcount = loopcount + 1

               else
                  WScript.Echo "FTP Failed (Stage 4)"
                  Stage = 99
               end if
          End if    

	   WScript.sleep 500
           WshShell.AppActivate strHost & " - PuTTY"
           WshShell.SendKeys "{ENTER}"

        'checks for a copy time out error and tries one more time.
        if TFTPERROR = "True" Then 
	   WScript.sleep 2500
           TFTP = "Only going to retry two times for each copy"
           Stage = 36
         end if
 
 'checks if file even exists.
        if TFTPERROR3 = "True" Then  
           msgtxt  = "No config file....." & vbCrLf & vbCrLf & "Check Salon# was Entered Correctly " & salonNum & vbCrLf & vbCrLf
           Stage = 99
         end if

         WScript.sleep 7500

  Case 444

	WshShell.SendKeys "reload~"  
	WScript.sleep 5000
        WshShell.SendKeys "~"
	WScript.sleep 500
	WshShell.SendKeys "n~"
	WScript.sleep 5000
        WshShell.SendKeys "~"  
        loopcount = 0
        txtdata = "User Access"
        Stage = 5

  case 5
     WScript.sleep 3000
     if IsAlive1 = "Pass" Then
	'WshShell.SendKeys "~"  
	'WScript.sleep 2000
        Stage = 6
        loopagain = True
        loopcount = 0

     else
        if loopcount < 60 then
	   loopagain = True
           loopcount = loopcount + 1
	   WshShell.SendKeys "~"
	   WScript.sleep 7500
 	else
           WScript.Echo "Reboot Error (Stage 5)"
           Stage = 99
        end if
     End if   

  case 6
         Stage = 1
         strpass = "123go!!!"
         struser = "guest"
	 ststage = 8
	 strcyc = 1
         loopagain = True
         loopcount = 0
	 txtdata = "name:"
         verdata2 = "rta-"
    
  case 7
     WScript.sleep 4000
     if IsAlive1 = "Pass" Then  
	 Stage = 8
     else
        if loopcount < 10 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "Failed to relogin after Reboot (Stage 7)"
           Stage = 99
        end if
     End if   

  case 8
     if salonNum = "Test" then
        msgbox "Press OK to test the Router." & vbCrLf & vbCrLf & "Close Putty window when complete" & vbCrLf 
	WScript.Quit
     end if

     WScript.sleep 500
     WshShell.SendKeys "config t~"
     WshShell.SendKeys "crypto key generate rsa modulus 1024~"
     WshShell.SendKeys "end~"
     WshShell.SendKeys "wr mem~"
     WScript.sleep 6000
     WshShell.SendKeys "y"
     WScript.sleep 1500
     WshShell.SendKeys "~"
     WScript.sleep 1500
     stage = 9
     txtdata = "Generating 1024 bit RSA key"
 
  case 9

     WScript.sleep 5000

     if IsAlive1 = "Pass" Then  
	 msgbox "Router configure for salon# " & salonNum & " is complete."
         WshShell.SendKeys "~"
         WScript.sleep 300
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
        WScript.Echo msgtxt & "Error please start over.",vbOKOnly, "Stage " & Stage 
	WshShell.AppActivate strHost & " - PuTTY"
        WshShell.SendKeys "exit~"
        WScript.sleep 2000
        WshShell.SendKeys "%{F4}"
        WScript.sleep 1000
        WshShell.SendKeys "{ENTER}"
	WScript.Quit

  case else
	loopagain = False
     
End Select

Loop While loopagain

msgbox "If you see this something didn't go right"
WshShell.AppActivate strHost & " - PuTTY"
        WshShell.SendKeys "exit~"
        WScript.sleep 2000
        WshShell.SendKeys "%{F4}"
        WScript.sleep 1000
        Wshshell.Sendkeys "~"
        WScript.Quit
