'datatable.ImportSheet "C:\Users\91983\OneDrive\Desktop\uft scripts\Common Floor\Import.xlsx",1,"Global"
'RunAction "Copy of Action1", oneIteration
'ExecuteFile "C:\Users\91983\OneDrive\Desktop\uft scripts\xpath.txt"

Services.StartTransaction "Selecting_City"
'opening the Browser and  select the city
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Aa).Set DataTable ("City1")
Browser("Property in India | Real").Page("Property in India | Real").WebEdit (L_Aa).Click

Browser("Property in India | Real").Page("Property in India | Real").WebList("MumbaiNavi MumbaiThaneRaigad").Select DataTable("City1")
Browser("Property in India | Real").Page("Property in India | Real").WebList("MumbaiNavi MumbaiThaneRaigad").Click

msgbox "city selected"

Services.EndTransaction "Selecting_City"

'Signup
If Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Ac).Exist (5) Then
	Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Ac).Click
	Reporter.ReportEvent micPass,"Confirm","profile is working"
End  If

Wait(5)
'Enter credentials from the external sheet
Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Ad).Click

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Ae).Set DataTable ("FullName")

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Af).Set DataTable ("Email")

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Ag).Set DataTable ("PhoneNo")

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Ah).SetSecure DataTable ("Password")

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Ai).Set DataTable("City1")

A = DataTable("You_are")
Select Case A
	Case "Owner"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Aj).Click
	Case "Builder"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Ak).Click
	Case "Broker"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_Al).Click
End Select

Browser("Property in India | Real").Page("Property in India | Real").WebButton(L_Am).Click

Wait(5)

'msgbox "you have been signed up"

Browser("Property in India | Real_2").Page("Property in India | Real").WebButton("Close").Click @@ script infofile_;_ZIP::ssf6.xml_;_

Wait (5)
'Login
If Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_AC).Exist (5) Then
	Browser("Property in India | Real").Page("Property in India | Real").WebElement(L_AC).Click
	Reporter.ReportEvent micPass,"Confirm","Profile is visible"
End  If
'Enter credentials  from external sheet for login
Browser("Property in India | Real_2").Page("Property in India | Real").WebElement("Log in to your account").Click @@ script infofile_;_ZIP::ssf8.xml_;_

Browser("Property in India | Real").Page("Property in India | Real").WebEdit(L_Af).Set DataTable ("PhoneNo")

Browser("Property in India | Real_2").Page("Property in India | Real").WebEdit("password").SetSecure "622106cffed85110d9078cbf3891f803ae861211" @@ script infofile_;_ZIP::ssf9.xml_;_

Browser("Property in India | Real").Page("Property in India | Real").WebButton(L_An).Click

msgbox "You have been Logged in"

Wait (5)
'Project selection
If Browser("Property in India | Real").Page("Property in India | Real").Link(L_Ao).Exist Then
	Browser("Property in India | Real").Page("Property in India | Real").Link(L_Ao).Click
	Reporter.ReportEvent micPass,"Confirm","Project Link is visible"
End If
Browser("Property in India | Real").Page("Property in India | Real").Link(L_Ao).Click

Browser("Property in India | Real").Page("Property in India | Real").Link(L_Ap).Click

Browser("Property in India | Real").Page("Completed Projects In").Check CheckPoint("Completed Projects In Pune | Commonfloor") @@ script infofile_;_ZIP::ssf11.xml_;_

Browser("Property in India | Real").Page("Property in India | Real").WebButton(L_Aq).Click

Browser("Property in India | Real").Page("Completed Projects In").WebElement("Price (Low to High)").Click

Browser("Property in India | Real").Page("Completed Projects In").WebButton("Contact").Click @@ script infofile_;_ZIP::ssf12.xml_;_

'Wait(5)
'Browser("Property in India | Real_2").Page("Completed Projects In").WebElement("XPath:=//div[@class='col-xs-12 col-sm-12 col-md-12']//div[contains(@class,'input_field open')]").Click
'Browser("Property in India | Real_2").Page("Completed Projects In").WebEdit("XPath:=//div[@class='col-xs-12 col-sm-12 col-md-12']//div[contains(@class,'input_field open')]").Set("XPath:=//div[@class='col-xs-12 col-sm-12 col-md-12']//div[contains(@class,'input_field open')]")

Browser("Property in India | Real").Page("Completed Projects In").WebButton("contact_2").Click @@ script infofile_;_ZIP::ssf13.xml_;_

T = Browser("Property in India | Real").Page("Completed Projects In").WebElement(L_Ar).GetROProperty("innertext") @@ script infofile_;_ZIP::ssf14.xml_;_
Reporter.ReportEvent micPass, "Confirm",T & " is visible "

Browser("Property in India | Real_2").Page("Completed Projects In").WebButton("Close").Click @@ script infofile_;_ZIP::ssf17.xml_;_

Browser("Property in India | Real_2").Page("Completed Projects In_2").Image(L_As).Click
