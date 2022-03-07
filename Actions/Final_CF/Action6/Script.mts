datatable.ImportSheet "C:\Users\Asif\Desktop\CF_FINAL\Import-Export.xls",1,"Global"
Set myBrowser=Browser("Property in India | Real").Page("Property in India | Real")
Services.StartTransaction "PostAds"
myBrowser.WebElement(abc).WaitProperty "visible", True, 2000
LstPropFor = datatable ("List_Proprety_For")
AvailFor = datatable ("Available_For")
Select Case LstPropFor
	Case "Sell"      'there are two options Sell And Rent
		myBrowser.WebElement(pp_sell).Click
		myBrowser.WebNumber("name:=price").Set datatable ("Price")
		myBrowser.WebElement(pp_price).Click
		If AvailFor="New"  Then
			myBrowser.WebElement(pp_new).Click
		Else
			myBrowser.WebElement(pp_resale).Click
		End If
	Case "Rent"
		myBrowser.WebElement(pp_rent).Click
		myBrowser.WebNumber("name:=price").Set datatable ("Rent_Price")
		myBrowser.WebNumber("name:=deposit_price").Set datatable ("Deposit_Price")
End Select
PropType = datatable ("Property_Type_2")
Select Case PropType    'to choose the property type
	Case "Apartment" 
		myBrowser.WebElement(pp_apar).Click
	Case "Builder Floor"
		myBrowser.WebElement(pp_bf).Click
	Case "Plot"
		myBrowser.WebElement(pp_plot).Click
	Case "House/Villa"
		myBrowser.WebElement(pp_hv).Click
End Select	
myBrowser.WebEdit(pp_date).Click
myBrowser.WebEdit(pp_date).Set datatable ("Date")
myBrowser.WebElement("XPath:=//div[contains(@class,'sc-1tonukx-0 hKjMKv cP Ell pR')]").Click

myBrowser.WebElement(pp_prop).WaitProperty "visible", True, 5000

Locali = "XPath:=//body[contains(@style,'width: 1903px; overflow: hidden;')]/div[@id='root']/div[contains(@class,'sc-guq2kf-0 kHvFRp')]/section[contains(@class,'sc-guq2kf-1 hWsjXB')]/div[contains(@class,'post__box dashed')]/div[contains(@class,'max__width')]/div[contains(@class,'sc-i5lwln-0 eCfkQt pF l0 t0 w100 h100 dFA jcC peN')]/div[contains(@class,'')]/div[contains(@class,'sc-i5lwln-6 hBoOTO oA')]/div[contains(@class,'sc-1cp1gbl-0 ikCpdx dF fwW')]/div[2]/div[1]/input[1]"
With Browser("Property in India | Real").Page("Post Property Ads for")
	.WebEdit("XPath:=//body[contains(@style,'width: 1903px; overflow: hidden;')]/div[@id='root']/div[contains(@class,'sc-guq2kf-0 kHvFRp')]/section[contains(@class,'sc-guq2kf-1 hWsjXB')]/div[contains(@class,'post__box dashed')]/div[contains(@class,'max__width')]/div[contains(@class,'sc-i5lwln-0 eCfkQt pF l0 t0 w100 h100 dFA jcC peN')]/div[contains(@class,'')]/div[contains(@class,'sc-i5lwln-6 hBoOTO oA')]/div[contains(@class,'sc-1cp1gbl-0 ikCpdx dF fwW')]/div[1]/div[1]/input[1]").Set datatable ("City")
	City = datatable ("City")
	Select Case City
		Case "Bangalore"
			.WebElement("XPath:=//span[contains(text(),'Bangalore')]").Click
			.WebEdit(Locali).Set datatable ("Locality")
			.WebElement("XPath:=//span[contains(text(),'Yelahanka New Town')]").Click
		Case "Chennai"
			.WebElement("XPath:=//span[contains(text(),'Chennai')]").Click
			.WebEdit(Locali).Set datatable ("Locality")
			.WebElement("XPath:=//span[contains(text(),'East Coast Road - ECR')]").Click
		Case "Hyderabad"
			.WebElement("XPath:=//span[contains(text(),'Hyderabad')]").Click
			.WebEdit(Locali).Set datatable ("Locality")
			.WebElement("XPath:=//span[contains(text(),'Kukatpally Industrial Estate')]").Click
	End  Select

	UTyp = datatable ("Unit_Type")
	Select Case UTyp
		Case "1RK"
			.WebElement("XPath:=//span[contains(text(),'1 RK')]").Click
		Case "1BHK"
			.WebElement("XPath:=//span[contains(text(),'1 BHK')]").Click
		Case "2BHK"
			.WebElement("XPath:=//span[contains(text(),'2 BHK')]").Click
		Case "3BHK"
			.WebElement("XPath:=//span[contains(text(),'3 BHK')]").Click
		Case "4+BHK"
			.WebElement("XPath:=//span[contains(text(),'4+ BHK')]").Click
	End  Select

	.WebNumber("XPath:=//body[contains(@style,'width: 1903px; overflow: hidden;')]/div[@id='root']/div[contains(@class,'sc-guq2kf-0 kHvFRp')]/section[contains(@class,'sc-guq2kf-1 hWsjXB')]/div[contains(@class,'post__box dashed')]/div[contains(@class,'max__width')]/div[contains(@class,'sc-i5lwln-0 eCfkQt pF l0 t0 w100 h100 dFA jcC peN')]/div[contains(@class,'')]/div[contains(@class,'sc-i5lwln-6 hBoOTO oA')]/div[contains(@class,'sc-1cp1gbl-0 ikCpdx dF fwW')]/div[5]/div[1]/div[1]/div[1]/input[1]").Set datatable ("Built_Up")
	.WebButton("XPath:=//button[contains(text(),'Submit')]").Click
End With


myBrowser.WebFile(pp_apj).Set "Appartment.jpg"



myBrowser.WebEdit(pp_add).Set datatable ("Additional_Info")
myBrowser.WebNumber("name:=property_on_floor").Set datatable ("Property_On_Floor")
myBrowser.WebNumber("name:=total_floors").Set datatable ("Total_Floors")

Furnishing = datatable ("Furnishing")
Select Case Furnishing
	Case "Semi Furnished" 
		myBrowser.WebElement(pp_sf).Click
	Case "Fully Furnished"  
		myBrowser.WebElement(pp_ff).Click
	Case  "Unfurnished"
		myBrowser.WebElement(pp_uf).Click
End Select

NoBR = datatable ("No_of_BathRooms")
Select Case NoBR
	Case 1
		myBrowser.WebElement(pp_cone).Click
	Case 2
		myBrowser.WebElement(pp_ctwo).Click
	Case 3
		myBrowser.WebElement(pp_cthree).Click
	Case 4
		myBrowser.WebElement(pp_cfour).Click
	Case "4+"
		myBrowser.WebElement(pp_cfourp).Click
End Select

myBrowser.WebElement("XPath:=//div[contains(@class,'sc-14cqhij-0 VVcsc cP pR')]").Click
myBrowser.WebEdit(pp_x).Set datatable ("No_of_Balconies")

NOB = datatable ("No_of_Balconies")
Select Case NOB
	Case 0
		myBrowser.WebElement("XPath:=//span[contains(text(),'0')]").Click
	Case 1		
		myBrowser.WebElement("XPath:=//span[@class='sc-1dhts6c-7 gJvvyB'][contains(text(),'1')]").Click
	Case 2 	
		myBrowser.WebElement("XPath:=//ul[@class='sc-1dhts6c-4 bXMhzy']//li[@class='sc-1dhts6c-5 fzdeLj pR cP']//span[@class='sc-1dhts6c-7 gJvvyB'][contains(text(),'2')]").Click
	Case 3
		myBrowser.WebElement("XPath:=//span[@class='sc-1dhts6c-7 gJvvyB'][contains(text(),'3')]").Click
	Case 4
		myBrowser.WebElement("XPath:=//body/div[@id='root']/div[@class='sc-guq2kf-0 kHvFRp']/section[@class='sc-guq2kf-1 hWsjXB']/div[@class='post__box dashed']/div[@class='max__width']/div[@class='sc-14cqhij-1 iwAxxk active']/div[@class='postad__boxrow']/div[@id='sD']/ul[@class='sc-1dhts6c-4 bXMhzy']/li[5]/span[1]").Click
	Case "4+" 	
		myBrowser.WebElement("XPath:=//span[@class='sc-1dhts6c-7 gJvvyB'][contains(text(),'4+')]").Click
End Select

myBrowser.WebEdit("XPath:=//body/div[@id='root']/div[@class='sc-guq2kf-0 kHvFRp']/section[@class='sc-guq2kf-1 hWsjXB']/div[@class='post__box dashed']/div[@class='max__width']/div[@class='sc-14cqhij-1 iwAxxk active']/div[2]/div[1]/input[1]").Click
myBrowser.WebElement(pp_club).Click
myBrowser.WebElement(pp_sec).Click

YouAre = datatable ("You_Are")
Select Case YouAre
	Case "Owner" 	
		myBrowser.WebElement(pp_own).Click
	Case "Agent" 		
		myBrowser.WebElement(pp_agen).Click
	Case "Builder" 	
		myBrowser.WebElement(pp_bul).Click
End Select

'Your Details------------------------------------------------------------------------

myBrowser.WebEdit(pp_name).Set datatable ("Name")
myBrowser.WebEdit(pp_mail).Set datatable ("Email")
myBrowser.WebEdit(pp_mno).Set datatable ("Mobile_No")
reporter.ReportEvent micDone, "Test Sucessful", "The user can post an Ad using post property."
Services.EndTransaction "PostAds"

