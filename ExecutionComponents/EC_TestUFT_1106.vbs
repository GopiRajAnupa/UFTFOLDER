Class clsEBS






Function  EC_TestUFT_1106 ()


	
	

	oControls_Web.CloseAllBrowsers BROWSER_NAME
		

	Systemutil.CloseProcessByName "rundll32.exe"
	systemutil.CloseProcessByName "javaw.exe"
	
	  oControls_Web.InvokeBrowser "https://oracleapps-test15.vmware.com/OA_HTML/AppsLocalLogin.jsp,BROWSER_NAME
	
	Call BrowserSynchronization()
      oSAPControls.SetSAPOKCode "/nva01"
		
aCreate_Sales_Order_InitialData=oSAPControls.GetHeaderScreenData ("Create Sales Order: Initial",Environment.Value("inputdata")) 
	
oSAPControls.SetSAPHeaderScreenData   aCreate_Sales_Order_InitialData
		
oSAPControls.SendSAPKey ENTER
		
aCreate_ZStandard_OrderData=oSAPControls.GetHeaderScreenData ("Create ZStandard Order:",Environment.Value("inputdata")) 
	
oSAPControls.SetSAPHeaderScreenData   aCreate_ZStandard_OrderData
		
oSAPControls.SendSAPKey ENTER
		
oSAPControls.SendSAPKey ENTER
		
oSAPControls.ClickSAPSaveButton  ()
		
sStatusMessage= oSAPControls.GetSAPStatusBarInfo  ("text")

	

	Call BrowserSynchronization()
	Reporter.EndFunction

End Function

Function GetValue(PropfilePath,ReqKey)

Dim oFS
Dim sPFSpec
Dim oTS
Set oFS = CreateObject( "Scripting.FileSystemObject" )
Set objFile1 = oFS.OpenTextFile( PropfilePath )

               do until objFile1.AtEndOfStream
                                  strLine= objFile1.ReadLine()
                                  strline1 = split(strLine,"=")
            If  strcomp(trim(strline1(0)),trim(ReqKey))=0 Then
                             Exit do
            End If
                                                           
               Loop
     GetValue = strline1(1) 
                
End Function


Function SetValue(PropfilePath,Keyname,KeyValue)

Const fsoForAppend = 8

Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream
Set objTextStream = objFSO.OpenTextFile(PropfilePath, fsoForAppend)

'Display the contents of the text file
objTextStream.WriteLine Keyname&" = "&KeyValue

'Close the file and clean up
objTextStream.Close
Set objTextStream = Nothing
Set objFSO = Nothing
SetValue =Array( Keyname,KeyValue)

End Function
End Class
