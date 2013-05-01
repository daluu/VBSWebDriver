'VBScript - WebDriver
'License: This script is distributed under the GNU General Public License 3.
'Author: henrytejera@gmail.com

Set Exception = New WebDriverException

Class WebDriverException
	
	Private sLastError

	''
	'Function: getError
	'Handling VBScript Run-Time Errors
	'
	'Parameters:
	'	obErr -  Err object
	'
	'Returns:
	'	Void
    Public Function getError(byRef objErr) 
    		sLastError = "Error: " & objErr.Number & vbCrLf & _ 
    					 "Error (Hex): " & Hex(objErr.Number)& vbCrLf & _
    					 "Source: " &  objErr.Source & vbCrLf & _
    					 "Description: " &  objErr.Description    	
			Call showLastError()
    		objErr.clear    		
    End Function    
    
	''
	'Function: setCommandsError
	'Handling WebDrive Commands Errors
	'
	'Parameters:
	'	sCode - String - Response Status Code
	'
	'Returns:
	'	Void
    Public Sub setCommandsError(byVal sCode) 
    		Dim sSummary : sSummary = ResponseStatusHandling.getResponseSummary(CInt(sCode))
    		sLastError = "Error code: " & sCode & vbCrLf & _ 
    					 "Error Summary: " & sSummary    					 
			Call showLastError()    		
    End Sub
    
    ''
	'Function: showLastError
    Public Sub showLastError()
    		MsgBox sLastError
    End Sub    
    
End Class