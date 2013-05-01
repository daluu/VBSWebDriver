'VBScript - WebDriver
'License: This script is distributed under the GNU General Public License 3.
'Author: henrytejera@gmail.com


''
'Function: Include
'Includes and evaluates the specified file.
'
'Parameters:
'    strFile - string 
Sub Include(ByVal strFile)
	On Error Resume Next
   	Set objFs = CreateObject("Scripting.FileSystemObject")
   	Set WshShell = CreateObject("WScript.Shell")
   	strFile = WshShell.ExpandEnvironmentStrings(strFile)
   	file = objFs.GetAbsolutePathName(strFile)
   	Set objFile = objFs.OpenTextFile(strFile)
   	strCode = objFile.ReadAll
   	objFile.Close
   	ExecuteGlobal(strCode)
   	
   	If Err.Number <> 0 Then
		Exception.getError(Err)						
	End If
End Sub

Include "JSON_2.0.4.vbs"
Include "WebElement.vbs"
Include "WebDriverException.vbs"
Include "WebDriverResponseStatus.vbs"

Class WebDriver
		
	Public sBaseURL
	Public sSessionID	
	Private objHTTP  	
	Private sBrowser	
	Public className	
	Public cssSelector	
	Public id 		
	Public name		
	Public linkText 	
	Public partialLinkText		
	Public tagName 		
	Public xpath	
		
	''
	'Function: Class_Initialize
	'Constructor	
	Private Sub Class_Initialize
		On Error Resume Next		
		Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
		initializeLocators()
		
		If Err.Number <> 0 Or IsObject(objHTTP)= False Then
			Exception.getError(Err)						
		End If		
	End Sub
	
	''
	'Function: initializeLocators
	'Initializes locator strategies
	Private Sub initializeLocators()
		className = "class name"
		cssSelector = "css selector" 
		id = "id"
		name ="name"
		linkText = "link text"
		partialLinkText = "partial link text"	
		tagName = "tag name"
		xpath = "xpath"		
	End Sub

	''
	'Function: connect
	'Create a new session. The server should attempt to create a session that most closely matches the desired capabilities.
	'
	'Parameters:
	'	sHost - string
	'	sPort - string
	' 	sBrowserName - string
	'   sVersion - string			
	Public Function connect(ByVal sHost, ByVal sPort,ByVal sBrowserName, ByVal sVersion)
		 On Error Resume Next
		 Dim sRequest
		 Dim objAllCaps : Set objAllCaps = jsObject()
		 
		 sBaseURL = "http://" & sHost & ":" & sPort & "/wd/hub/session"
	
		 
		Set objAllCaps("desiredCapabilities") = jsObject()
		objAllCaps("desiredCapabilities")("javascriptEnabled") = "true"
		objAllCaps("desiredCapabilities")("nativeEvents") = "false"
		objAllCaps("desiredCapabilities")("browserName") = sBrowserName
		objAllCaps("desiredCapabilities")("version") = sVersion
				 		 
		sSessionID = getSession(executePost(sBaseURL,objAllCaps.jsString))	
		Set objAllCaps = Nothing
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
	End Function
	
	''
	'Function: close
	'Close the current session
	Public Sub close()
		On Error Resume Next
	    Dim sRequest : sRequest = sElementRequest						
	    
	    If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
	End Sub	
	
	''
	'Function: executePost
	'Sends an HTTP POST request to Remote WebDriver server.
	'
	'Parameters:
	'	sRequest - string
	'	sArgs - string
	'
	'Returns:
	'	String
	Public Function executePost(ByVal sRequest,ByVal sArgs)
		On Error Resume Next
		Dim sResponse
		
		With objHTTP
			.open "POST", sRequest, False 
			.setRequestHeader "Content-Type","application/json;charset=UTF-8"
			.send sArgs	
			sResponse = .responseText						
		End With					
		ExecutePost = sResponse
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
	End Function
	
	''
	'Function: executeGet
	'Sends an HTTP GET request to Remote WebDriver server.
	'
	'Parameters:
	'	sRequest - string
	'	sArgs - string
	'
	'Returns:
	'	server response - String	
	Public Function executeGet(ByVal sRequest)
		On Error Resume Next
		Dim sResponse
		
		With objHTTP
			.open "GET", sRequest, False 
			.setRequestHeader "Content-Type","application/json;charset=UTF-8"			
			.send	
			sResponse = .responseText						
		End With							
		ExecuteGet= sResponse
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If		
	End Function
	
	''
	'Function handleResponse
	'Function analyses status attribute of the response.
	'For some statuses it throws exception (for example NoSuchElementException)
	'
	'Parameters:
	'	sResponse - String 
	Public Sub handleResponse(ByVal sResponse)
		Dim sStatusCode
		
		If 	sResponse <> "" Then
			Dim parser : Set parser = jsonParser()	
			sStatusCode = parser.getProperty(sResponse, "status",False)			
			
			If sStatusCode <> 0 Then
				Exception.setCommandsError(sStatusCode)
			End If
			
			Set parser = Nothing
		End If	
				
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If		
	End Sub
	
	''
	'Function: getSession
	'Get session id
	'
	'Parameters:
	'	sResponseText - string
	'
	'Returns:
	'	Session id - String	
	Private Function getSession(ByVal sResponseText) 
		On Error Resume Next
		Dim objRegExpr, colMatches
		
		Set objRegExpr = New RegExp
		objRegExpr.Pattern = "RequestURI=/wd/hub/session/[0-9]{13}"
		objRegExpr.Global = True
		objRegExpr.IgnoreCase = True
		
		Set colMatches = objRegExpr.Execute(sResponseText)
		If colMatches.Count <> 0 Then
			GetSession = Mid(colMatches.Item(0),28,13)
		End If	
		
		Set colMatches = Nothing
		Set objRegExpr = Nothing
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If		
	End Function
	
	''
	'Function: closeWindow
	'Get session id
	'
	'Parameters:
	'	sResponseText - string
	'
	'Returns:
	'	Session id - String		
    Public Sub closeWindow()
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/window"		        
        Call executeGet(sRequest)
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If	                       		            
    End Sub

	''
	'Function: getWindowHandle
	'Retrieve the current window handle.
	'
	'Returns:
	'	The current window handle - String		
    Public Function getWindowHandle()
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/window_handle"		        
        getWindowHandle = executeGet(sRequest)
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If	                       		            
    End Function    

	''
	'Function: executeScript
	'Inject a snippet of JavaScript into the page for execution in the context of the currently selected frame. 
	'The executed script is assumed to be synchronous and the result of evaluating the script is returned to the client.
	'
	'Parameters:
	'	sScript - String - The script to execute.
	'	aArgs - Array - The script arguments.
	'
	'Returns:
	'	Result of evaluating the script is returned to the client.		
    Public Function executeScript(ByVal sScript,ByVal aArgs)
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/execute"		        
        Dim objArgs : Set objArgs = jsObject()
        
        
		objArgs("script")= sScript
		objArgs("args")=  Array(aArgs)		
		gexecuteScript = executePost(sRequest,objArgs.jsString)
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If	                       		            
    End Function    
	
	''
	'Function: navigateTo
	'Navigate to a new URL
	'
	'Parameters:
	'	sURL - string - The URL to navigate to
	'
	'Returns:
	'	Void
	Public Sub navigateTo(sURL)
		On Error Resume Next
		Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/url"
        
        Set objArgs = jsObject()
		objArgs("url")= sURL
        
        Call executePost(sRequest,objArgs.jsString)       
         
        Set objArgs = Nothing 
        If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
    End Sub

	''
	'Function: getCurrentUrl
	'Retrieve the URL of the current page
	'
	'Returns:
	'	The current URL - String
	Public Function getCurrentUrl()
		On Error Resume Next
		Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/url"
		Dim parser : Set parser = jsonParser()		
		Dim sURL : sURL = executeGet(sRequest)
		
		getCurrentUrL = parser.getProperty(sURL, "value",False)			
		Set parser = Nothing
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
    End Function	

	''
	'Function: getTitle
	'Get the current page title
	'
	'Returns:
	'	The current page title - String
	Public Function getTitle()
		On Error Resume Next
		Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/title"		
		Dim parser : Set parser = jsonParser()		
		Dim sTitle : sTitle = executeGet(sRequest)
		
		getTitle = parser.getProperty(sTitle,"value",False)
		Set parser = Nothing
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
    End Function     
 
	''
	'Function: getAlertText
	'Gets the text of the currently displayed JavaScript alert(), confirm(), or prompt() dialog.
	'
	'Returns:
	'	The text of the currently displayed alert - String
    Public Function getAlertText() 
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/alert_text"		
		Dim parser : Set parser = jsonParser()		
		Dim sAlertText : sAlertText = executePost(sRequest,Null)                
        		
		getAlertText = parser.getProperty(sAlertText,"value",False)				
		Set parser = Nothing	
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If				
    End Function
    
	''
	'Function: acceptAlert
	'Accepts the currently displayed alert dialog. Usually, this is equivalent to clicking on the 'OK' button in the dialog.
	'
	'Returns:
	'	Void
    Public Sub acceptAlert()
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/accept_alert"		        
		Call executeGet(sRequest) 
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If                       		            
    End Sub

	''
	'Function: dismissAlert
	'Dismisses the currently displayed alert dialog. For confirm() and prompt() dialogs, this is equivalent 
	'to clicking the 'Cancel' button. For alert() dialogs, this is equivalent to clicking the 'OK' button.
	'
	'Returns:
	'	Void
    Public Sub dismissAlert()
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/dismiss_alert"		
		Call executePost(sRequest,Null) 
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If                          		            
    End Sub
        
	''
	'Function: getScreenshot
	'Take a screenshot of the current page.
	'
	'Returns:
	'	The screenshot as a base64 encoded PNG - String
    Public Function getScreenshot() 
    	On Error Resume Next
        Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/screenshot"	
		Dim parser : Set parser = jsonParser()		
		Dim sScreenshot : sScreenshot = executeGet(sRequest)                
        		
		getScreenshot = parser.getProperty(sScreenshot,"value",False)
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If       			
    End Function         
	
	''
	'Function: getScreenshot
	'Refresh the current page.
	'
	'Returns:
	'	Void
	Public Sub refresh()
		On Error Resume Next
		Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/refresh"
        Call ExecutePost(sRequest,Null) 
        
        If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If                 
    End Sub    	

	''
	'Function: findElementBy
	'Search for an element on the page, starting from the document root.
	'
	'Parameters:
	'	sLocatorStrategy - String - The locator strategy to use.
	'	sValue - String - The search target.
	'
	'Returns:
	'	A WebElement JSON object for the located element - Object
	Public Function findElementBy(ByVal sLocatorStrategy, ByVal sValue)
		On Error Resume Next
		Dim sRequest : sRequest = sBaseURL & "/" & sSessionID & "/element"	
		Dim objAllCaps : Set objAllCaps = jsObject()
		Dim parser : Set parser = jsonParser()		
		Dim oElement : Set oElement = New WebElement
		Dim sResponse
				
		objAllCaps("using") = sLocatorStrategy
		objAllCaps("value") = sValue
						 		 
		sResponse = executePost(sRequest,objAllCaps.jsString)			
		
		oElement.Init me,parser.getProperty(sResponse,"value","ELEMENT")
		
		Set objAllCaps = Nothing		
		Set parser = Nothing
		
		Set findElementBy = oElement
		
        If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 		
	End Function
	
End Class	