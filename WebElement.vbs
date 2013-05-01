'VBScript - WebDriver
'License: This script is distributed under the GNU General Public License 3.
'Author: henrytejera@gmail.com

Class WebElement

	Public sElementID
	Private objDriver	
	Private sElementType	
	Private sElementRequest
	
	''
	'Function: Init
	'Constructor
	'
	'Parameters:
	'	objWebDriver - WebDriver Object
	'	sValue - String
	Public Sub Init(ByRef objWebDriver,ByVal sValue)
		On Error Resume Next
		Set objDriver = objWebDriver
				
		sElementID = sValue
		sElementRequest = objDriver.sBaseURL & "/" & objDriver.sSessionID & "/element/"	
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 				
	End Sub

	''
	'Function: sendKeys
	'Send a sequence of key strokes to an element.
	'
	'Parameters:
	'	sValue - String - The sequence of keys to type.	
	Public Sub sendKeys(ByVal sValue) 
		On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/value"	
		Dim objElement : Set objElement = jsObject()
				
		objElement("value") = Array(sValue)						 		 		
		Dim sResponse : sResponse = objDriver.executePost(sRequest,objElement.jsString)
		objDriver.handleResponse sResponse	
		Set objElement = Nothing    
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	    
    End Sub
        
	''
	'Function: click
	'Click any mouse button (at the coordinates set by the last moveto command). 
	'Note that calling this command after calling buttondown and before calling button up 
	'(or any out-of-order interactions sequence) will yield undefined behaviour).	
    Public Sub click() 
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/click"							
		Dim sResponse : sResponse = objDriver.executePost(sRequest,Null)
        objDriver.handleResponse sResponse	
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	 
	End Sub    
	
	''
	'Function: click
	'Submit a FORM element. The submit command may also be applied to any element that is a descendant of a FORM element.	
	Public Sub submit()
		On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/submit"
        Dim sResponse : sResponse = objDriver.executePost(sRequest,Null)
        objDriver.handleResponse sResponse	
        
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	        				
	End Sub

	''
	'Function: getText
	'Returns the visible text for the element.
	'
	'Returns:
	'	Returns the visible text for the element - String
    Public Function getText() 
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/text"			
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)
		
		objDriver.handleResponse sResponse	        
		getText = parser.getProperty(sResponse,"value",False)				
		Set parser = Nothing		
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 		
    End Function		
    
	''
	'Function: getName
	'Query for an element's tag name.
	'
	'Returns:
	'	The element's tag name, as a lowercase string. - String    
    Public Function getName()
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/name"			
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)                
		objDriver.handleResponse sResponse
        		
		getName = parser.getProperty(sResponse,"value",False)				
		Set parser = Nothing

		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 			
	End Function	    

	''
	'Function: clear
	'Clear a TEXTAREA or text INPUT element's value.
    Public Sub clear() 
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/clear"			
        Dim sResponse : sResponse = objDriver.executePost(sRequest,Null)
        objDriver.handleResponse sResponse

		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	        		
    End Sub
    
	''
	'Function: getAttributeName	 
	'Get the value of an element's attribute.
	'
	'Returns:
	'	The value of the attribute, or null if it is not set on the element - String
    Public Function getAttributeName()
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/attribute/name"			
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)   
		objDriver.handleResponse sResponse                       
                			
		getAttributeName = parser.getProperty(sResponse,"value","ELEMENT")				
		Set parser = Nothing       

		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If
    End Function
    
	''
	'Function: isEnabled
	'Determine if an element is currently enabled.
	'
	'Returns:
	'	Whether the element is enabled - String	     
    Public Function isEnabled()
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/enabled"			
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)
		objDriver.handleResponse sResponse                       
		    		
		isEnabled = parser.getProperty(sResponse,"value","ELEMENT")
		
		Set parser = Nothing  

		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 			
    End Function

	''
	'Function: isDisplayed
	'Determine if an element is currently displayed.
	'
	'Returns:
	'	Whether the element is displayed - String	      
    Public Function isDisplayed()
      	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/displayed"			
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)     
		objDriver.handleResponse sResponse
        		
		isDisplayed = parser.getProperty(sResponse,"value","ELEMENT")				
		Set parser = Nothing  
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	
    End Function
       
	''
	'Function: getCssProperty
	'Query the value of an element's computed CSS property. The CSS property to query should be specified using the CSS property name, 
	'not the JavaScript property name (e.g. background-color instead of backgroundColor).
	'
	'Parameters:
	'	sPropertyName - String - Property name
	'
	'Returns:
	'	The value of the specified CSS property - String       
    Public Function getCssProperty(ByVal sPropertyName)
    	On Error Resume Next
        Dim sRequest : sRequest = sElementRequest & sElementID & "/css/"& sPropertyName
		Dim parser : Set parser = jsonParser()		
		Dim sResponse : sResponse = objDriver.executeGet(sRequest)     
		objDriver.handleResponse sResponse          
        		
		getCssProperty = parser.getProperty(sResponse,"value","ELEMENT")				
		Set parser = Nothing  
		
		If Err.Number <> 0 Then
			Exception.getError(Err)						
		End If 	
    End Function     
            
End Class 

