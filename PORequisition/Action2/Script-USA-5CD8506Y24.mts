'============================================================================================================================================================ 
' Purpose:  Create Sales Order
' Created by:  Quynh Ly
' Created Date:  Aug 12
'-------------------------------------------------------------------------------------------- 
' Modified Date:
' Modified by:
'============================================================================================================================================================ 
 
' Declare variable
 Dim Vn_Cnt, Vn_Data_Cnt
 Dim Vc_Return_SalesOrder,Vc_Sales_Order, Vc_Business_Unit, Vc_Customer, Vc_Contact, Vc_Purchase_Order, Vc_Order_Type, Vc_Bill_To_Customer, Vc_Bill_To_Account, Vc_Ship_To_Customer, Vc_Ship_To_Address
 Dim Vc_Utimate_Destination, Vc_Rig_Name, Vc_Order_Manager, Vc_Warehouse, Vc_Item, Vc_Quantity, Vc_Currency
   
' Initialie variables
 Vn_Cnt = 1
 Vn_Data_Cnt = Datatable.GetSheet("Create Sales Order").GetRowCount
 Datatable.GetSheet("Create Sales Order").SetCurrentRow Vn_Cnt

 DO 
 
	 'Home
	 Home()
	 wait 1
	
	'Select Order Management
	 OracleMainMenu("Order Management")
	  
	 OrderManagement("Order Management")
	  
	' Initialize variables
	Vc_Org=Datatable("Org", "Create Sales Order")
	Vc_Business_Unit=Datatable("Business_Unit", "Create Sales Order")
	Vc_Customer=Datatable("Customer", "Create Sales Order")
	Vc_Contact=datatable("Contact", "Create Sales Order")
	Vc_Purchase_Order=Datatable("Purchase_Order", "Create Sales Order")
	Vc_Order_Type=Datatable("Order_Type","Create Sales Order")
	Vc_Bill_To_Customer=Datatable("Bill_To_Customer", "Create Sales Order")
	Vc_Bill_To_Account=Datatable("Bill_To_Account", "Create Sales Order")
	Vc_Ship_To_Customer=Datatable("Ship_To_Customer", "Create Sales Order")
	Vc_Ship_To_Address=datatable("Ship_To_Address", "Create Sales Order")
	Vc_Utimate_Destination=datatable("Utimate_Destination", "Create Sales Order")
	Vc_Rig_Name=Datatable("Rig_Name", "Create Sales Order")
	Vc_Order_Manager=Datatable("Order_Manager", "Create Sales Order")
	Vc_Warehouse=datatable("Warehouse", "Create Sales Order")
	Vc_Item=datatable("Item", "Create Sales Order")
	Vc_Quantity=datatable("Quantity", "Create Sales Order")
	Vc_Payment_Term= Datatable("Payment_Term", "Create Sales Order")
	Vc_Currency= Datatable("Currency", "Create Sales Order")
	
	'Create Sales Order
	 CreateSalesOrder Vc_Org, Vc_Business_Unit, Vc_Customer, Vc_Contact, Vc_Purchase_Order, Vc_Order_Type, Vc_Bill_To_Customer, Vc_Bill_To_Account, Vc_Ship_To_Customer, Vc_Ship_To_Address,Vc_Utimate_Destination, Vc_Rig_Name, Vc_Order_Manager, Vc_Warehouse, Vc_Item, Vc_Quantity, Vc_Payment_Term,Vc_Currency
	 wait 2
	
	' Capture SO and View the Order in the Fullfillment View
	 Vc_Return_SalesOrder= CapturedCreatedSalesOrder(Vc_Sales_Order)
	 Datatable("Created_Order", "Create Sales Order") = Vc_Return_SalesOrder
	 wait 1
	    
	' Increment counter
	 Vn_Cnt = Vn_Cnt + 1
	 Datatable.GetSheet("Create Sales Order").SetNextRow

Loop Until Vn_Cnt > Vn_Data_Cnt - 1
wait 1

' Home
 Home()

'Logoff
 Logoff()
 
 @@ hightlight id_;_Browser("Manage Orders - Order").Page("Oracle Applications").WebElement("Good afternoon, Quynh-Nhu")_;_script infofile_;_ssf210.xml_;_
