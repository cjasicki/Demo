'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Name: BrotherMFPSetup.vbs                                              '
'  By: Chad Jasicki                                                       '
'  Version 02222012   							  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Function: Checks and cycles through a range of IP addess. Then it      '
'            Prompts user for salon number. Then verifies the input. 	  '  
'            Then it formats the salon Number and compiles those variables'
'            into the salon name. Then it runs BrSet.exe tool and set     '
'	     printers Node Name, SNMP Get and Set Community Names,        '
'	     Password. Then logs result in a log.txt file. 		  '
'                                                                         '
'	     Must have DHCP server with 10.10.10.x network pool           '
'                                                                         '
'                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declare variables

Dim ip4
Dim IP3 
Dim pingResult
Dim sCurPath
Dim RS1
Dim clcr

'----------------------------------------------------------------
'  Node Name Format(numSalon)                                   '
'----------------------------------------------------------------
Function SalonFormat(numSalon)
    Dim intZeroes
    Dim z

     'Set counter to format length to 8 characters
     'and add leading 0's to number
    verSalNum = Len(numSalon)
    If verSalNum <> 8 Then
        intZeroes = 8 - verSalNum
        For z = 1 to intZeroes
            numSalon = "0" & numSalon
        Next
    End If

     'Return value
    SalonFormat = numSalon

End Function


'----------------------------------------------------------------
'  Delete setup files                                           '
'----------------------------------------------------------------
'Sub Delete()
'objFSO.DeleteFile (sCurPath & "\*.tmp")
'End Sub
'
'----------------------------------------------------------------
'  Creates/appends results to a log file.                       '
'----------------------------------------------------------------
Sub Log()

dim objOutFile
Dim sWorkingFileName

Const FOR_APPENDING = 8
sWorkingFileName = sCurPath & "\log.txt"
Set objFSO1 = Wscript.CreateObject("Scripting.FileSystemObject")

If objFSO1.FileExists(sWorkingFileName) Then
  Set objOutFile = objFSO1.OpenTextFile(sWorkingFileName, FOR_APPENDING)
Else
  Set objOutFile = objFSO1.CreateTextFile(sWorkingFileName)
  objOutFile.WriteLine "  Node Name      Date      Time       Results"
End If

if rc1 = 0 and rc2 = 0 and rc11 = 0 and rc22 = 0 then    
rcw = "successful"
else
rcw = "error: " & rc11 & " - " & rc22 & " - " & rc1 & " - " & rc2 
end if

objOutFile.WriteLine formSalonNum &"  "& date() &" " & time() & "  " & rcw
objOUTFile.Close

End Sub

'----------------------------------------------------------------
'  Ping Command - checks to see if printer is pinable.          '
'----------------------------------------------------------------
Sub PingTR()
pingResult = False
Const OpenAsASCII = 0 
Const FailIfNotExist = 0 
Const ForReading =  1 
Dim objShell, objFSO, sTempFile, fFile
Set objFSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("Wscript.Shell")
objName = objFSO.GetTempName
sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
pingResult = False
objShell.Run "cmd /c ping -n 1 -w 15 -l 1 "& IP3 & ip4 &">" & objName, 0, True    
'Set objTextFile = objFSO.OpenTextFile(objName, 1)
Set fFile = objFSO.OpenTextFile(objName, ForReading, FailIfNotExist, OpenAsASCII) 
	Select Case InStr(fFile.ReadAll, "Reply") 
         Case 0
            pingResult = False 

         Case Else
            pingResult = True

	End Select 
 
fFile.Close 
objFSO.DeleteFile (sCurPath & "\*.tmp")
End Sub

'-------------------------------------------------------------------
'  IP Address 				                            '
'  The variable "IP3" is the first three octets of the network      '  
'  The veriable "IP4" is the range the scrip will look for printers ' 
'-------------------------------------------------------------------
IP3 = "10.10.10."
StrIP = 0
clcr = 0
FMFP = 0

For ip4 = 100 to 111

   PingTR()

   If pingResult = True Then
     StrIP = (ip4 -1) + 1 
     FMFP = FMFP + 1
   end if

Next
	If StrIP = 0 Then
           msgbox "No Printer Found Please Check that the Nextwork Cable is Plugged In." 
        Else
           msgbox "Found " & FMFP & " Printers. Press OK to Configure the Printer/s"
        end if

For ip4 = 100 to StrIP
                RS1 = 7
		Do while RS1 = 7
		salonNum = inputbox("Printer IP Address: "& IP3 & ip4 & chr(13) & chr(13) & chr(13) & "Enter the salon number for this pinter:", "Salon Number")
		  If salonNum = "" then
                        salonNum = "error"
                        RR12 = msgbox ("you must enter a salon#",,"error")
		  end If 
		formSalonNum = "BRN1-" & SalonFormat((salonNum))
                  If salonNum >< "error" then
		      RS1 = msgbox ("Is this the correct Node Name for the Printer?"& chr(13) & chr(13) & "                " & formSalonNum,32 + 4,"Printer/Node Name")
		  End If                 
		Loop
		
		Do
		i = 0
		intSS = inputbox("Enter the appropriate salon SSID for the printer: " & chr(13) & chr(13) & _
					"1 For rgs-promenade-hwres"  & chr(13) & _
					"2 For rgs-supercuts-hwres"  & chr(13) & _
					"3 For rgs-smartstyle-hwres"  & chr(13) & _
					"4 For rgs-mastercuts-hwres"  & chr(13) & _
					"5 For rgs-regis_salon-hwres"  & chr(13) & _
					"6 For rgs-sassoon-hwres press"  & chr(13) & _
					"7 For rgs-@@@@@@-hwres press" &  chr(13) & _
					"8 For rgs-******-hwres press"  & chr(13) & chr(13), "SSID")	
		
		Select Case intSS
		Case ""
                       i = 5
                       RR12 = msgbox ("you must pick one",,"SSID")
		Case 1
			intSSID = "rgs-promenade-hwres"
		Case 2
			intSSID = "rgs-supercuts-hwres"
		Case 3
			intSSID = "rgs-smartstyle-hwres"
		Case 4
			intSSID = "rgs-mastercuts-hwres"
		Case 5
			intSSID = "rgs-regis_salon-hwres"
		Case 6
			intSSID = "rgs-sassoon-hwres"
		Case 7
			intSSID = "rgs-@@@@@@@@-hwres"
		Case 8
			intSSID = "rgs-********-hwres"
		
		Case Else
			msgbox "Invalid Selection.  Please Try Again", vbCritical, "Invalid Selection"
			i = 5
		End Select

		RS121 = 7
                if i >< 5 then
		RS121 = msgbox ("you picked:  " & intSSID & chr(13) &" Is that Correct?",32 + 4,"SSID")
		end if
                if RS121 = 7 then
		i = 5
		end if

	Loop While i = 5	

		getComm = "public"
		newGetComm = "ryJCldYvmQ"
		SetComm = "tKxQqqTqmv"
		nodename = formSalonNum
		Channel = "11"
		WPSKey = "Xv!c07tj9zAb"
		intssg = 1
		loopagain = True
		Set WshShell = WScript.CreateObject("WScript.Shell")

        do

	Select Case intssg
	Case 1
		rc11 = wshShell.Run ("%comspec% /c %CD%\RegisDepTool\smtpset "& IP3 & ip4 & " public -Regis_Salons" ,5,True)
                'WScript.Sleep 6000

                  If rc11 <> 0 Then
                     RS11 = MsgBox ("smtpset Command Failed - Error Code: " & rc11 & chr(13) & chr(13) & "Press Retry..." & chr(13) & chr(13) & "If Retry didn't work reset the Priner to factory default" ,16 + 5 + 256,"BrSet failed")
                  Else
                     intssg = intssg + 1
                  End If     

        Case 2
                Shutt = 0
                do While Shutt = 0    'Need to add some error checking for the this loop statment
                   WScript.Sleep 500
  	           PingTR()
                   if pingResult = False Then
                      Shutt = 1
                   End if

                Loop
                WScript.Sleep 8000
                intssg = intssg + 1

        Case 3
                Shutt = 0
                do While Shutt = 0   'Need to add some error checking for the this loop statment
                   WScript.Sleep 500
  	           PingTR()
                   if pingResult = True Then
                      Shutt = 1
                   End if

                Loop
                WScript.Sleep 5000
                intssg = intssg + 1
 
        Case 4

    		rc22 = wshShell.Run ("%comspec% /c %CD%\RegisDepTool\pushAddr -IP "& IP3 & ip4 & " public" ,5,True)
                WScript.Sleep 2000

                  If rc22 <> 0 Then
                     RS22 = MsgBox ("PushAddr Command Failed - Error Code: " & rc22 & chr(13) & chr(13) & "Press Retry..." & chr(13) & chr(13) & "If Retry didn't work reset the Priner to factory default" ,16 + 5 + 256,"BrSet failed")
                  Else
                     intssg = intssg + 1
                  End If     


        Case 5
                rc1 = wshShell.Run ("%comspec% /c %CD%\RegisDepTool\brWlan.exe "& IP3 & ip4 & " " & intSSID & " " & Channel & " " & WPSKey & " -node " & nodename ,5,True)
                WScript.Sleep 6000

                  If rc1 <> 0 Then
                     RS1 = MsgBox ("BrWlan Command Failed - Error Code: " & rc1 & chr(13) & chr(13) & "Press Retry..." & chr(13) & chr(13) & "If Retry didn't work reset the Priner to factory default" ,16 + 5 + 256,"BrSet failed")
                  Else
                     intssg = intssg + 1
                  End If     

        Case 6
		rc2 = wshShell.Run ("%comspec% /c %CD%\RegisDepTool\brsetnc.exe " & IP3 & ip4 & " " & "public" & " -cpw KcRg!!! -cgc " & newGetComm & " -csc " & SetComm & " -cnn " & formSalonNum ,5,True)             
                WScript.Sleep 5000
                  If rc2 <> 0 Then
	             RTP = wshShell.Run ("%comspec% /c %CD%\RegisDepTool\brsetnc.exe " & IP3 & ip4 & " force_defaultGC",5,True) 
                     WScript.Sleep 1500
                     RS2 = MsgBox ("BrSetNC Command Failed - Error Code: " & rc2 & chr(13) & chr(13) & "Press Retry..." & chr(13) & chr(13) & "If Retry didn't work reset the Priner to factory default" ,16 + 5 + 256,"BrSet failed")
                  Else
                     intssg = intssg + 1
                  End If     
        Case 7
		'MsgBox "Printer Setup for Salon# " & formSalonNum & " is Completed",0,"MPF Printer Settup Completed - " & rc2
		loopagain = False
                clcr = clcr + 1
		log()

	Case Else
		msgbox "? Start Over!!!!!!??????"
		loopagain = False
	End Select		

        Loop While loopagain

  ' end if 		
next
msgbox "Script Completed" & chr(13) & chr(13) & "Total Printer: " & clcr,64,""