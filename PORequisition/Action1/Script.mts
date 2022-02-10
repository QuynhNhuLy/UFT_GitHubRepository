'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Purpose:  OSP Requisition Creation
'--------------------------------------------------------------------------------------------------------
' Created by:  Quynh Ly
' Created date:  Oct 2020
'--------------------------------------------------------------------------------------------------------
' Modified by:
' Modified date:
'----------------------------------------------------------------------------------------------------------------------------------------------------------
     
'Declare variables   
Dim Vc_Line_Type, Vc_Item, Vc_Quantity, Vc_UOM, Vc_Price, Vc_Currency, Vc_Requester, Vc_Request_date, Vc_Agreement_Type
Dim Vc_Supplier, Vc_Supplier_Site, Vc_Supplier_Contact, Vc_Deliver_To, Vc_Destination_Type, Vc_Project, Vc_COA, Vc_Percentage, Vc_Amount, Vc_Note_To_Supplier
Dim Vn_Cnt, Vn_Data_Cnt, Saved_Inventory_Org
    
 
'Initialize variables
Vn_Cnt = 1
Vn_Data_Cnt = Datatable.GetSheet("OSP Requisition Creation").GetRowCount
Datatable.GetSheet("OSP Requisition Creation").SetCurrentRow Vn_Cnt
Saved_Inventory_Org = ""
Saved_Legacy_PO = ""
First_Time_PO = "True"
'Vc_Current_Date = Month(Date)&"/"&Day(Date)+1&"/"&Year(Date)
Vc_Current_Date = "24/1/2022"
Vc_Saved_Error_Msg =""
Reset_PO = "False"

DO
    Vc_Shipping_Delivery_Requester=Datatable("Shipping_Delivery_Requester", "OSP Requisition Creation")
	Vc_Shipping_Delivery_Deliver_To=Datatable("Shipping_Delivery_Deliver_To", "OSP Requisition Creation")
	
    If Saved_Inventory_Org <> RTRIM(LTRIM(Datatable("Inventory_Org", "OSP Requisition Creation"))) Then
               
       'Home 
		Home()  
		wait (.2)
		 
		OracleMainMenu("Procurement")
		wait (.2)
		
		Procurement("Purchase Requisitions")
		wait (.2)
		
        'Set_References Vc_Shipping_Delivery_Requester, Vc_Shipping_Delivery_Deliver_To
        wait (.2)
       
        Saved_Inventory_Org = RTRIM(LTRIM(Datatable("Inventory_Org", "OSP Requisition Creation")))
    End If
          
  ' Enter Requisition Lines
	Vc_Line_Type=Datatable("Line_Type", "OSP Requisition Creation")
	Vc_Item=Datatable("Item", "OSP Requisition Creation")
'	Vc_Quantity=Datatable("Quantity", "OSP Requisition Creation")
'	Vc_UOM=Datatable("UOM", "OSP Requisition Creation")
	Vc_Price=Datatable("Price", "OSP Requisition Creation")
	Vc_Currency=Datatable("Currency", "OSP Requisition Creation")
	Vc_Delivery_Requester=Datatable("Delivery_Requester", "OSP Requisition Creation")
    Vc_Request_Date=Datatable("Request_Date", "OSP Requisition Creation")
    
    If CDate(Vc_Request_Date) < CDate(Vc_Current_Date) Then
       Vc_Request_Date = Vc_Current_Date 
       'Datatable("Error", "OSP Requisition Creation") = "Requested delivery date is in the past. Current Date is used"
       'Msgbox Vc_Request_Date 
    End If    
    
'	Vc_Agreement_Type=Datatable("Agreement_Type", "OSP Requisition Creation")
	Vc_Supplier=RTRIM(LTRIM(Datatable("Supplier", "OSP Requisition Creation")))
	Vc_Supplier_Site=RTRIM(LTRIM(Datatable("Supplier_Site", "OSP Requisition Creation")))
	Vc_Supplier_Contact=RTRIM(LTRIM(Datatable("Supplier_Contact", "OSP Requisition Creation")))
	Vc_Delivery_Deliver_To=RTRIM(LTRIM(Datatable("Delivery_Deliver_To", "OSP Requisition Creation")))
	Vc_Destination_Type=RTRIM(LTRIM(Datatable("Destination_Type", "OSP Requisition Creation")))
'	Vc_Project=Datatable("Project", "OSP Requisition Creation")
'	Vc_COA=Datatable("COA", "OSP Requisition Creation")
'	Vc_Percentage=Datatable("Percentage", "OSP Requisition Creation")
'	Vc_Amount=Datatable("Amount", "OSP Requisition Creation")
    Vc_Note_To_Supplier=RTRIM(LTRIM(Datatable("Note_To_Supplier", "OSP Requisition Creation")))
    Vc_Work_Order = RTRIM(LTRIM(Datatable("Work_Order",  "OSP Requisition Creation")))
    Vc_Operation_Seq = RTRIM(LTRIM(Datatable("Operation_Seq", "OSP Requisition Creation")))
    Vc_Suggested_Buyer = RTRIM(LTRIM(Datatable("Suggested_Buyer", "OSP Requisition Creation")))
    Vc_Legacy_PO=RTRIM(LTRIM(Datatable("Legacy_PO", "OSP Requisition Creation")))
    Vc_Legacy_PO_Line=RTRIM(LTRIM(Datatable("Legacy_PO_Line", "OSP Requisition Creation")))
    Vc_Expenditure_Type = RTRIM(LTRIM(Datatable("Expenditure_Type", "OSP Requisition Creation")))
    Vc_Expenditure_Organization = RTRIM(LTRIM(Datatable("Expenditure_Organization", "OSP Requisition Creation")))

    
	' Create Requisition
	Vc_Error = OSP_Create_Requisition(Vc_Line_Type, Vc_Item, Vc_Price, Vc_Currency, Vc_Delivery_Requester, Vc_Request_Date, Vc_Supplier, Vc_Supplier_Site, Vc_Supplier_Contact, Vc_Delivery_Deliver_To, Vc_Destination_Type, Vc_Work_Order,Vc_Operation_Seq, Vc_Suggested_Buyer, Vc_Note_To_Supplier, Vc_Legacy_PO, First_Time_PO,Saved_Legacy_PO,Vc_Legacy_PO_Line, Vc_Expenditure_Type, Vc_Expenditure_Organization, Vc_Saved_Error_Msg, Reset_PO) 
	If RTRIM(LTRIM(Vc_Error <> "")) Then
		Reset_PO = "True"
    Else
        Reset_PO = "False"
	End If
	
	' Write error	
	Vc_Saved_Error_Msg  = Vc_Error
	Datatable("Error", "OSP Requisition Creation") = Vc_Error
	
'	Msgbox Vc_Error
	
	'Increment counters
	First_Time_PO = "False"
	
	Vn_Cnt = Vn_Cnt + 1
	Datatable.GetSheet("OSP Requisition Creation").SetNextRow
	
Loop until Vn_Cnt > 1
wait 3

'===============================================================================
'Submit and Capture Last PO
'===============================================================================

	 		
'Create a browser
 Set OBrowser=description.Create
 OBrowser("micclass").value="Browser"
 OBrowser("title").value=".*" 
	
'Create a page
 Set OPage=description.Create
 OPage("micclass").value="Page"
 OPage("title").value=".*" 

 Set OImage=description.Create
 OImage("micclass").value="Image"
 OImage("image type").value="Image Link"	  
 OImage("html tag").value="IMG"
 OImage("alt").value="Shopping Cart"
 
	'Create a button
	Set OButton=description.Create
	OButton("micclass").value="WebButton"
	OButton("type").value="submit"
	OButton("name").value="View PDF"
		
  ' Select Shopping Card
  Browser(OBrowser).Page(OPage).Image(OImage).Click
  wait 1
		
  ' Submit
  Submit_Bttn()
  wait 1
		
  ' Capture Requisition
  Vc_Requisition_No = CaptureRequisition(Vc_Requisition)
  wait 1
		
  Datatable("OSP_Requisition", "OSP Requisition Creation") =  Vc_Requisition_No
  
 Browser(OBrowser).Page(OPage).WebButton(OButton).Click 
      
 Datatable.GetSheet("OSP Requisition Creation").SetPrevRow
  
  wait 1
		
  
