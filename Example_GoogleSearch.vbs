''
'Simple Google search test
'VBScript - WebDriver

Sub Include(ByVal strFile)
   Set objFs = CreateObject("Scripting.FileSystemObject")
   Set WshShell = CreateObject("WScript.Shell")
   strFile = WshShell.ExpandEnvironmentStrings(strFile)
   file = objFs.GetAbsolutePathName(strFile)
   Set objFile = objFs.OpenTextFile(strFile)
   strCode = objFile.ReadAll
   objFile.Close
   ExecuteGlobal(strCode)
End Sub

Include "WebDriver.vbs"


Set Driver = New WebDriver
	Driver.connect "127.0.0.1","4444","internet explorer", ""
	Driver.navigateTo "http://www.google.com"	
	MsgBox "Retrieve the URL of the current page: " & Driver.getCurrentUrl()
	Driver.executeScript "alert('Simple Google search test')",""

Set Element  = Driver.findElementBy(Driver.name,"q")
	Element.sendKeys "VBScript"
	Element.submit
	MsgBox "Element's tag name: " & Element.getName()
	MsgBox "Element attribute name: " & Element.getAttributeName
	MsgBox "Element is enabled?: " & Element.isEnabled
	MsgBox "Element is displayed?: " & Element.isDisplayed	
	MsgBox "Element CSS Color property: " & Element.getCssProperty("color")
		
	