'VBScript - WebDriver
'License: This script is distributed under the GNU General Public License 3.
'Author: henrytejera@gmail.com

Set ResponseStatusHandling = New WebDriverResponseStatus

Class WebDriverResponseStatus 

	Private ResponseStatusCodes
	
	''
	'Function: Class_Initialize
	'Constructor    
    Public Sub Class_Initialize
    	On Error Resume Next
    	Set ResponseStatusCodes = createObject("Scripting.Dictionary")
    	Call populateResponseStatus
    	
    	If Err.Number <> 0 Or IsObject(ResponseStatusCodes) = False Then
    		MsgBox Err.Number
    	End If    	    	
    End Sub	
    
    ''
    'Function: populateResponseStatus      
    Private Sub populateResponseStatus()
    	'The command executed successfully.
    	ResponseStatusCodes.Add 0,"Success"
    	'An element could not be located on the page using the given search parameters.
    	ResponseStatusCodes.Add 7,"NoSuchElement"
    	'A request to switch to a frame could not be satisfied because the frame could not be found.
    	ResponseStatusCodes.Add 8,"NoSuchFrame"
    	'The requested resource could not be found, or a request was received using an HTTP method that is not supported by the mapped resource
    	ResponseStatusCodes.Add 9,"UnknownCommand"
    	'An element command failed because the referenced element is no longer attached to the DOM.
    	ResponseStatusCodes.Add 10,"StaleElementReference"
    	'An element command could not be completed because the element is not visible on the page.
    	ResponseStatusCodes.Add 11,"ElementNotVisible"
    	'An element command could not be completed because the element is in an invalid state (e.g. attempting to click a disabled element).
    	ResponseStatusCodes.Add 12,"InvalidElementState"
    	'An unknown server-side error occurred while processing the command.
    	ResponseStatusCodes.Add 13,"UnknownError"
    	'An attempt was made to select an element that cannot be selected.
    	ResponseStatusCodes.Add 15,"ElementIsNotSelectable"
    	'An error occurred while executing user supplied JavaScript.
    	ResponseStatusCodes.Add 17,"JavaScriptError"
    	'An error occurred while searching for an element by XPath.
    	ResponseStatusCodes.Add 19,"XPathLookupError"
    	'An operation did not complete before its timeout expired.
    	ResponseStatusCodes.Add 21,"Timeout"
    	'A request to switch to a different window could not be satisfied because the window could not be found.
    	ResponseStatusCodes.Add 23,"NoSuchWindow"
    	'An illegal attempt was made to set a cookie under a different domain than the current page.
    	ResponseStatusCodes.Add 24,"InvalidCookieDomain"
    	'A request to set a cookie's value could not be satisfied.
    	ResponseStatusCodes.Add 25,"UnableToSetCookie"
    	'A modal dialog was open, blocking this operation
    	ResponseStatusCodes.Add 26,"UnexpectedAlertOpen"
    	'An attempt was made to operate on a modal dialog when one was not open.
    	ResponseStatusCodes.Add 27,"NoAlertOpenError"
    	'A script did not complete before its timeout expired.
    	ResponseStatusCodes.Add 28,"ScriptTimeout"   	 
    	'The coordinates provided to an interactions operation are invalid.   	    	    	    	
		ResponseStatusCodes.Add 29,"InvalidElementCoordinates"    	    	    	    	    	    	
		'IME was not available.
		ResponseStatusCodes.Add 30,"IMENotAvailable" 
		'An IME engine could not be started.
		ResponseStatusCodes.Add 31,"IMEEngineActivationFailed" 
		'Argument was an invalid selector (e.g. XPath/CSS).			    	    	    	    	    			
		ResponseStatusCodes.Add 32,"InvalidSelector"		
    End Sub 
    
    ''
    'Function: getResponseSummary
    '
    'Parameters:
    '	sCode - String - Response Status Code
    '
    'Returns:
    '	Response Status summary - String
    Public Function getResponseSummary(ByVal sCode)
    	getResponseSummary = ResponseStatusCodes.Item(Int(sCode))
    End Function
        
End Class
