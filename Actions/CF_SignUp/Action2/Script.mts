'----------------Location of the external Excel sheet...............
datatable.ImportSheet "C:\Users\Administrator\Desktop\Capgemini\Import23.xls",1,"Global"
datatable.ExportSheet "C:\Users\Administrator\Desktop\Capgemini\Import23.xls","Global","Global"
ExecuteFile("C:\Users\Administrator\Desktop\Capgemini\Constants.txt")
 '----------------------Link of the Application..........................
'URL="https://www.commonfloor.com"
'SystemUtil.Run "Firefox",URL

Services.StartTransaction "Searching_Home" '--------------Transaction Starts
 '------------------------Homepage of the Application.............................
'Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Ta).WaitProperty "visible", True, 7000
'Browser("Property in India | Real").Page("Property in India | Real").WebEdit(S_Tb).Set DataTable("S_City")      'Selection of the City for Home..................
'Browser("Property in India | Real").Page("Property in India | Real").WebEdit(S_Tb).Click
'Browser("Property in India | Real").Page("Property in India | Real").WebList("MumbaiNavi MumbaiThaneRaigad").Select DataTable("S_City")
Wait(7)
Browser("Property in India | Real").Page("Property in India | Real").WebButton(S_Sc).Click

'----------------Type of Home either Rent or Buy.......................	
AB = datatable ("Buy_Rent")
If AB= "Buy" Then
	Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tc).Click
	Else
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Td).Click
End If

Wait(3)
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(S_Te).Click

'-------------Selection of Builder or Project...................
BuildProp = datatable ("Builder")
Select Case BuildProp
	Case "Pristine Properties"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tf).Click
	Case "Kohinoor Group"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tg).Click
	Case "Ganga Florentina"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Th).Click
	Case "Paranjape Blue Ridge"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Ti).Click
End Select

Services.EndTransaction "Searching_Home"     '--------End Transaction---------

'------------------Filters that we insert as per ourwish or dreamhome......................
Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tj).Click

Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tk).WaitProperty "visible", True, 5000
'----------------Selection of type of property like Villa, Appartment....................
Aj = datatable ("Property_Type")
Select Case Aj
	Case "Villa"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tl).Click
	Case "Appartment"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tm).Click
End Select

Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tn).Click
Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_To).WaitProperty "visible", True, 5000

'------------------Select the rooms how many we want.............................
Ay = datatable ("BHK")
Select Case Ay
	Case "2BHK"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tp).Click
	Case "3BHK"
		Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tq).Click
End Select

'------------------Select the area of the property................................
Sat = "XPath:=//body/div[@id='filters']/div[@id='mobilefiltershow']/div[@id='allfilter']/div[@class='col-xs-12 col-sm-12']/div[@class='row']/div[@class='col-xs-12 col-cm-2 pdltrgnone']/div[@class='custom-dd']/ul[@class='custom-dd-minmax buy-rent-filter-class']/li/div[@class='custom-minmax clearfix']/div[1]/input[1]"
Mat = "XPath:=//body/div[@id='filters']/div[@id='mobilefiltershow']/div[@id='allfilter']/div[@class='col-xs-12 col-sm-12']/div[@class='row']/div[@class='col-xs-12 col-cm-2 pdltrgnone']/div[@class='custom-dd']/ul[@class='custom-dd-minmax buy-rent-filter-class']/li/div[@class='custom-minmax clearfix']/div[2]/input[1]"
Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tr).Click
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(Sat).Click
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(Sat).Set datatable ("CarpetArea_Min")
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(Mat).Click
Browser("Property in India | Real").Page("Property in India | Real").WebEdit(Mat).Set datatable ("CarpetArea_Max")

Browser("Property in India | Real").Page("Property in India | Real").Link(S_Ts).Click
Browser("Property in India | Real").Page("Property in India | Real").WebElement(S_Tt).WaitProperty "visible", True, 5000

 '-------------Generating the report--------------
AZ=Ay
AX=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tu).GetROProperty ("innertext")
A=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tv).GetROProperty("innertext")
B=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tw).GetROProperty("innertext")
C=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tx).GetROProperty("innertext")
If Instr(AZ,AX)=0 Then
	Reporter.ReportEvent micPass, "Search Result","This result is the subset of filtered search"
	Reporter.ReportEvent micPass, "Confirm","The choosen property with rent price " &A& " with available carpet area of " &B& "	is " &C
	Else
		Reporter.ReportEvent micFail, "Search Result","This result is not the subset of filtered search"
End If

L=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tu).GetROProperty("innertext")
M=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tz).GetROProperty("innertext")
N=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tab).GetROProperty("innertext")
DataTable.Value("Appartment_Name")=L
DataTable.Value("Parking")=M
DataTable.Value("Property_On")=N
DataTable.Value("Address")=O

'-----------Contacting details of the owner for the belonging property..........................................
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tah, S_Tax).Click
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tac).Click
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebButton(S_Tad).Click
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement("Thank You").Check CheckPoint("Thank You")

Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tae).Click
Wait(4)
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Taf).WaitProperty "visible", True, 5000
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tag).Click
O=Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tag).GetROProperty("innertext")
Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").WebElement(S_Tai).WaitProperty "visible", True, 5000
DataTable.Value("Address")=O

Browser("Rent 3 BHK Semi-Furnished_3").Page("Rent 3 BHK Semi-Furnished").Link("XPath:=//a[@class='postpropbtn cf-tracking-enabled']").Click









