''----------------Location of the external Excel sheet...............
'datatable.ImportSheet "C:\Users\Administrator\Desktop\Capgemini\Import23.xls",1,"Global"
'datatable.ExportSheet "C:\Users\Administrator\Desktop\Capgemini\Import23.xls","Global","Global"
'ExecuteFile("C:\Users\Administrator\Desktop\Capgemini\Constants.txt")
' '----------------------Link of the Application..........................
''URL="https://www.commonfloor.com"
''SystemUtil.Run "Firefox",URL

Services.StartTransaction "Searching_Home" '--------------Transaction Starts
 '------------------------Homepage of the Application.............................
'Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Ta).WaitProperty "visible", True, 7000
'Browser("Property in India | Real").Page("Property in India | Real").WebEdit(S_Tb).Set DataTable("S_City")      'Selection of the City for Home..................
'Browser("Property in India | Real").Page("Property in India | Real").WebEdit(S_Tb).Click
'Browser("Property in India | Real").Page("Property in India | Real").WebList("MumbaiNavi MumbaiThaneRaigad").Select DataTable("S_City")
Wait(7)
With Browser("Property in India | Real").Page("Property in India | Real")
	.WebButton(S_Sc).Click

	'----------------Type of Home either Rent or Buy.......................	
	AB = datatable ("Buy_Rent")
	If AB= "Buy" Then
		.WebElement(S_Tc).Click
		Else
			.WebElement(S_Td).Click
	End If

	Wait(3)
	.WebEdit(S_Te).Click

	'-------------Selection of Builder or Project...................
	BuildProp = datatable ("Builder")
	Select Case BuildProp
		Case "Pristine Properties"
			.WebElement(S_Tf).Click
		Case "Kohinoor Group"
			.WebElement(S_Tg).Click
		Case "Ganga Florentina"
			.WebElement(S_Th).Click
		Case "Paranjape Blue Ridge"
			.WebElement(S_Ti).Click
	End Select

	Services.EndTransaction "Searching_Home"     '--------End Transaction---------

	'------------------Filters that we insert as per ourwish or dreamhome......................
	.WebElement(S_Tj).Click

	.WebElement(S_Tk).WaitProperty "visible", True, 5000
	'----------------Selection of type of property like Villa, Appartment....................
	Aj = datatable ("Property_Type")
	Select Case Aj
		Case "Villa"
			.WebElement(S_Tl).Click
		Case "Appartment"
			.WebElement(S_Tm).Click
	End Select

	.WebElement(S_Tn).Click
	.WebElement(S_To).WaitProperty "visible", True, 5000

	'------------------Select the rooms how many we want.............................
	Ay = datatable ("BHK")
	Select Case Ay
		Case "2BHK"
			.WebElement(S_Tp).Click
		Case "3BHK"
			.WebElement(S_Tq).Click
	End Select

	'------------------Select the area of the property................................
	Sat = "XPath:=//body/div[@id='filters']/div[@id='mobilefiltershow']/div[@id='allfilter']/div[@class='col-xs-12 col-sm-12']/div[@class='row']/div[@class='col-xs-12 col-cm-2 pdltrgnone']/div[@class='custom-dd']/ul[@class='custom-dd-minmax buy-rent-filter-class']/li/div[@class='custom-minmax clearfix']/div[1]/input[1]"
	Mat = "XPath:=//body/div[@id='filters']/div[@id='mobilefiltershow']/div[@id='allfilter']/div[@class='col-xs-12 col-sm-12']/div[@class='row']/div[@class='col-xs-12 col-cm-2 pdltrgnone']/div[@class='custom-dd']/ul[@class='custom-dd-minmax buy-rent-filter-class']/li/div[@class='custom-minmax clearfix']/div[2]/input[1]"
	.WebElement(S_Tr).Click
	.WebEdit(Sat).Click
	.WebEdit(Sat).Set datatable ("CarpetArea_Min")
	.WebEdit(Mat).Click
	.WebEdit(Mat).Set datatable ("CarpetArea_Max")
wait(5)
	'.Link(S_Ts).Click
End With
Browser("Property in India | Real").Page("86 Properties for rent").Link("Semi Furnished 3BHK Apartment").Click @@ script infofile_;_ZIP::ssf1.xml_;_
'Browser("72 Properties for rent").Page("Rent 3 BHK Semi-Furnished").WebElement("3BHK Apartment for Rent").Click @@ script infofile_;_ZIP::ssf2.xml_;_
	'.WebElement(S_Tt).WaitProperty "visible", True, 5000


 '-------------Generating the report--------------
AZ=Ay
With Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished")
	AX=.WebElement(S_Tu).GetROProperty ("innertext")
	A=.WebElement(S_Tv).GetROProperty("innertext")
	B=.WebElement(S_Tw).GetROProperty("innertext")
	C=.WebElement(S_Tx).GetROProperty("innertext")
	If Instr(AZ,AX)=0 Then
		Reporter.ReportEvent micPass, "Search Result","This result is the subset of filtered search"
		Reporter.ReportEvent micPass, "Confirm","The choosen property with rent price " &A& " with available carpet area of " &B& "	is " &C
		Else
			Reporter.ReportEvent micFail, "Search Result","This result is not the subset of filtered search"
	End If

	L=.WebElement(S_Tu).GetROProperty("innertext")
	M=.WebElement(S_Tz).GetROProperty("innertext")
	N=.WebElement(S_Tab).GetROProperty("innertext")
	DataTable.Value("Appartment_Name")=L
	DataTable.Value("Parking")=M
	DataTable.Value("Property_On")=N
	DataTable.Value("Address")=O

	'-----------Contacting details of the owner for the belonging property..........................................
	.WebElement(S_Tah, S_Tax).Click
	.WebElement(S_Tac).Click
	.WebButton(S_Tad).Click
	'.WebElement("Thank You").Check CheckPoint("Thank You")

	.WebElement(S_Tae).Click
	Wait(4)
	.WebElement(S_Taf).WaitProperty "visible", True, 5000
	.WebElement(S_Tag).Click
	O=.WebElement(S_Tag).GetROProperty("innertext")
	.WebElement(S_Tai).WaitProperty "visible", True, 5000
	DataTable.Value("Address")=O
	
	
Browser("Property in India | Real").Page("Property in India | Real").Link("Ashish").Click @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Property in India | Real").Page("Property in India | Real").WebElement("Log Out").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Property in India | Real").Page("Property in India | Real").Sync
Browser("Property in India | Real").Refresh @@ hightlight id_;_1509098_;_script infofile_;_ZIP::ssf5.xml_;_

	.Link("XPath:=//a[@class='postpropbtn cf-tracking-enabled']").Click
End With











