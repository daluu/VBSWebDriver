Sub Include(ByVal strFile)
   Set objFs = CreateObject("Scripting.FileSystemObject")
   Set WshShell = CreateObject("WScript.Shell")
   strFile = WshShell.ExpandEnvironmentStrings(strFile)
   file = objFs.GetAbsolutePathName(strFile)      
   Set objFile = objFs.OpenTextFile(File)
   strCode = objFile.ReadAll
   objFile.Close
   ExecuteGlobal(strCode)
End Sub

Sub killProcess(ByVal strProcessName)
    set colProcesses = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_Process Where Name='" & strProcessName & "'")
    If colProcesses.count <> 0 then
        For Each objProcess in colProcesses
            objProcess.Terminate()
        Next
    End if
End Sub 

Include "WebDriver.vbs"

Dim Driver 
Dim sSUT : sSUT = "http://htejera.users.sourceforge.net/vbswebdriver/demo/"
Dim sBrowser : sBrowser = "internet explorer"

sub setup		
	Set Driver = New WebDriver
	Driver.connect "127.0.0.1","4444",sBrowser, ""
end sub

sub Teardown()
	Set Driver = Nothing
	Select Case sBrowser 
		Case "internet explorer"
			 killProcess "iexplore.exe"
		Case "firefox"
			 killProcess "firefox.exe"
	End Select		 
end sub

sub Test_GetTitle()       	
	Assert.IsSomething Driver, "object Driver was not created" 
					
	Driver.NavigateTo sSUT
	title = Driver.getTitle()			
	Assert.IsEqual title, "jQuery accordion form with validation", "getTitle"
end sub

sub Test_GetCurrentURL()       	
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	url = Driver.getCurrentUrl()		
	Assert.IsEqual url, sSUT, "Get Current URL"
end Sub

Sub Test_FindElementByName() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim combo : Set combo = Driver.findElementBy(Driver.name,"recordPurchaseTimeFrameID") 
	Assert.IsSomething combo, "object Element by name was not created" 
End Sub
		
		
Sub Test_FindElementById() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim combo : Set combo = Driver.findElementBy(Driver.id,"recordPurchaseTimeFrameID") 
	Assert.IsSomething combo, "object Element by id was not created" 
End Sub

Sub Test_FindElementByXpath() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim input : Set input = Driver.findElementBy(Driver.xpath,"//*[@id='prod_name']") 
	Assert.IsSomething input, "object Element by xpath was not created" 
End Sub

Sub Test_FindElementByLinkText() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim link : Set link = Driver.findElementBy(Driver.linkText,"Next") 
	Assert.IsSomething link, "object Element by link text was not created" 
End Sub

Sub Test_ElementGetType() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim combo : Set combo = Driver.findElementBy(Driver.name,"recordPurchaseState") 
	Assert.IsSomething combo, "object Element by name was not created" 
		
	Assert.IsEqual combo.getName,"select","getName error"
End Sub

Sub Test_ElementGetAttributeName() 
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim combo : Set combo = Driver.findElementBy(Driver.name,"recordPurchaseState") 
	Assert.IsSomething combo, "object Element by name was not created" 
		
	Assert.IsEqual combo.getAttributeName,"recordPurchaseState","getAttributeName error"
End Sub

Sub Test_isDisplayed()
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim radio : Set radio = Driver.findElementBy(Driver.name,"recordPurchaseMetRealtor") 
	Assert.IsSomething radio, "object Element by name was not created" 
		
	Dim sResult : sResult = radio.isDisplayed	
	Assert.IsTrue sResult, "isDisplayed error"
End Sub

Sub Test_isEnabled()
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim combo : Set combo = Driver.findElementBy(Driver.name,"recordPurchasePropertyTypeID") 
	Assert.IsSomething combo, "object Element by name was not created" 
	
	Dim sResult : sResult = combo.isEnabled	
	Assert.IsTrue sResult, "isEnabled error"	
End Sub


Sub Test_getWindowHandle()
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Dim sResult : sResult = Driver.getWindowHandle()
	
	Assert.Trace "Window Handle response"
  	Assert.Remark sResult
	Assert.NotEqual sResult,"",sResult			
End Sub

Sub Test_ExecuteScriptAndGetText()
	Assert.IsSomething Driver, "object Driver was not created" 
	
	Driver.NavigateTo sSUT
	Driver.executeScript "document.write('<h1 id=""text"">Hello World!</h1>');",""
	Dim Element : Set Element = Driver.findElementBy(Driver.xpath,"//*[@id='text']") 
	Assert.IsSomething Element, "object Element by name was not created" 
	
	Dim sText : sText = Element.getText()	
	Assert.IsEqual sText,"Hello World!"	
End Sub

Sub Test_completeForm1() 
	Assert.IsSomething Driver, "object Driver was not created" 
	Driver.NavigateTo sSUT
	
	'Are you currently working with a real estate agent?
	Dim radio : Set radio = Driver.findElementBy(Driver.xpath,"//*[@id='sf1']/div/fieldset/input[2]") 
	Assert.IsSomething radio, "object Element by name was not created" 
	radio.click
	Set radio = Nothing
	
	'When would you like to move?
	Dim combo1 : Set combo1 = Driver.findElementBy(Driver.xpath,"//*[@id='recordPurchaseTimeFrameID']/option[2]") 
	Assert.IsSomething combo1, "object Element by name was not created" 
	combo1.click
	Set combo1 = Nothing
	
	'Purchase price range
	Dim combo2 : Set combo2 = Driver.findElementBy(Driver.xpath,"//*[@id='recordPurchasePriceRangeID']/option[6]") 
	Assert.IsSomething combo2, "object Element by name was not created" 
	combo2.click
	Set combo2 = Nothing
	
	'State
	Dim combo3 : Set combo3 = Driver.findElementBy(Driver.xpath,"//*[@id='recordPurchaseState']/option[6]") 
	Assert.IsSomething combo3, "object Element by name was not created" 
	combo3.click
	Set combo3 = Nothing
	
	'Desired property type
	Dim city : Set city = Driver.findElementBy(Driver.xpath,"//*[@id='recordCityName']") 
	Assert.IsSomething city, "object Element by name was not created" 
	city.sendKeys "Montevideo"
	Set city = Nothing
	
	
	'Desired property type
	Dim combo4 : Set combo4 = Driver.findElementBy(Driver.xpath,"//*[@id='recordPurchasePropertyTypeID']/option[3]") 
	Assert.IsSomething combo4, "object Element by name was not created" 
	combo4.click
	Set combo4 = Nothing
	
	'Next
	Dim button : Set button = Driver.findElementBy(Driver.name,"formNext1") 
	Assert.IsSomething button, "object Element by name was not created" 
	button.click
	Set button = Nothing
	
	'Verification
	Dim input : Set input = Driver.findElementBy(Driver.name,"recordPropertyAddress1") 
	Assert.IsSomething input , "object Element by name was not created" 
	Set input = Nothing
End Sub
