'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Purpose:  Login testing for test
'--------------------------------------------------------------------------------------------------------
' Created by:  Quynh Ly
' Created date:  Oct 2020
'--------------------------------------------------------------------------------------------------------
' Modified by:
' Modified date:
'----------------------------------------------------------------------------------------------------------------------------------------------------------
 
''  Declare variable
''Dim  Vc_Customer_Rating, Vc_Pricing_Segment, Vc_Precedence, Vc_Start_Date
'' SystemUtil.Run "C:\Program Files\Google\Chrome\Application\chrome.exe","https://ekbn-dev4.login.us6.oraclecloud.com/"
' 
'    'Create a browser
'	Set OBrowser=description.Create
'	OBrowser("micclass").value="Browser"
'	OBrowser("title").value=".*" 
'	    
'	'Create a page
'	Set OPage=description.Create
'	OPage("micclass").value="Page"
'	OPage("title").value="Sign In" 
'	
'	'Create a button
'	Set OButton=description.Create
'	OButton("micclass").value="WebButton"
'	OButton("name").value="Sign in"
'	OButton("html tag").value="SPAN"
'	

' Single sign on
SingleSignOn Datatable("Browser", dtlocalsheet), Datatable("Instance",  "Login")
wait 1

Login()
 @@ hightlight id_;_10000000_;_script infofile_;_ssf5.xml_;_

