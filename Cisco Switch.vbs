' config setup for Cisco 8 port switch

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

struser = "cisco"
strpass = "cisco"
loopcount = 0
Stage = 88
bx = "No"
verdata2 = "ap#"
ststage = 2
TFTP = "%Error opening tftp"
aces = "% Login invalid" 
txtdata = "#exit"

set WshShell = WScript.createObject ("WScript.Shell")
WshShell.run "C:\putty.exe -load Cisco_Switch"
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
	txtdata = "User Name:"
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
	WshShell.SendKeys struser & "~"
        WScript.sleep 1500
	WshShell.SendKeys strpass & "~"
	WshShell.SendKeys "n"
	loopagain = True
	txtdata = "switch"
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
	strpass = "pass2"
	struser = "user2"
         verdata2 = "rta-" 
	 txtdata = "% Authentication failed"
         stage = 1
     end if
															
  case 3
      	WScript.sleep 500
 	WshShell.SendKeys "~"
  	WshShell.SendKeys "conf t~"
	WshShell.SendKeys "vlan database~"
	WshShell.SendKeys "vlan 20~"
	WshShell.SendKeys "exit~"
	WshShell.SendKeys "int fa8~"
	WshShell.SendKeys "no macro auto smartport~"
	WshShell.SendKeys "switchport mode access~"
	WshShell.SendKeys "switchport access vlan 20~"
	WshShell.SendKeys "exit~"
	WshShell.SendKeys "int vlan 20~"
	WshShell.SendKeys "ip address dhcp~"
	WshShell.SendKeys "y"
	WshShell.SendKeys "end~"
	WshShell.SendKeys "copy tftp://10.1.97.62/switch_" & salonNum & ".txt startup-config~"
	WshShell.SendKeys "y"
        strpass = ""
        struser = ""
        txtdata = "bytes copied in"	        
        Stage = 4
	
         WScript.Quit

   case 4

     WScript.sleep 5000
    
     logfile()

     if IsAlive1 = "Pass" Then  
        WScript.sleep 800
	WshShell.SendKeys "reload~"  
	WScript.sleep 800
	WshShell.SendKeys "y~"  
	WScript.sleep 800
	WshShell.SendKeys "y~"  
	WScript.sleep 800
        Stage = 5
        txtdata = "Resetting local unit"

     else
        if loopcount < 30 then
	   loopagain = True
           loopcount = loopcount + 1  
 	else
           WScript.Echo "TFTP Failed. Ensure file exists. (Stage 4)"
           Stage = 99
        end if
     End if    

     if TFTPERROR = "True" Then  
         Stage = 3   
	 WScript.sleep 10000   
         TFTP = "nogoanddonothing123321"
     end if
   
  case 5

     WScript.sleep 3000

     logfile()

     if IsAlive1 = "Pass" Then  
	WshShell.SendKeys "~"  
	WScript.sleep 800

     else
      	WshShell.SendKeys "no~" 
	WScript.sleep 2000
	WshShell.SendKeys "~"
     End if   

     txtdata = "User Name:"
     Stage = 6
     Loopagain = True
     Loopcount = 0

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
	 msgbox "Check prompt for swtich-" & salonNum & "-10. If Prompt is good... Access point is done" 

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
