'Welcome. Please read carefully.
'Only change what is between the lines right below these instructions.

'advisorNumber is your advisor number. Make sure you put it between the quotation marks.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim advisorNumber
advisorNumber = "74247"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function loadConfigFile()
  'Find the path name of this script
  strPath = Wscript.scriptFullname
  'Create File System Object
  Set objFSO = createObject("Scripting.FileSystemObject")
  'Create object for this script's file
  Set objFile = objFSO.getFile(strPath)
  'Get the folder that this script is located in
  strFolder = objFSO.getParentFolderName(objFile)
  'Create an Excel Object
  Set objExcel = createObject("Excel.Application")
  'Open the Rebill Pricing Excel File
  Set objWorkbook = objExcel.workBooks.open(strFolder & "\Rebill Pricing.xlsx")

  Dim rowcount
  '-----------------------------Customer Pricing-----------------------------
  'Load the sheet and store the data
  Set objCustomerPricingSheet = objWorkbook.worksheets("Customer Pricing")
  rowcount = objCustomerPricingSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve customerPricingData(4, i - 2)
    customerPricingData(0, i - 2) = objCustomerPricingSheet.cells(i, 2)
    customerPricingData(1, i - 2) = objCustomerPricingSheet.cells(i, 3)
    customerPricingData(2, i - 2) = objCustomerPricingSheet.cells(i, 4)
    customerPricingData(3, i - 2) = objCustomerPricingSheet.cells(i, 5)
  Next

  'End of data, clear memory
  Set objCustomerPricingSheet = Nothing

  '-----------------------------Vehicle Pricing-----------------------------
  'Load the sheet and store the data
  Set objVehiclePricingSheet = objWorkbook.worksheets("Vehicle Pricing")
  rowcount = objVehiclePricingSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve vehiclePricingData(4, i - 2)
    vehiclePricingData(0, i - 2) = objVehiclePricingSheet.cells(i, 1)
    vehiclePricingData(1, i - 2) = objVehiclePricingSheet.cells(i, 2)
    vehiclePricingData(2, i - 2) = objVehiclePricingSheet.cells(i, 3)
    vehiclePricingData(3, i - 2) = objVehiclePricingSheet.cells(i, 4)
  Next

  'End of data, clear memory
  Set objVehiclePricingSheet = Nothing

  '-----------------------------Customer Number Change-----------------------------
  'Load the sheet and store the data
  Set objCustomerNumberChangeSheet = objWorkbook.worksheets("Customer Number Change")
  rowcount = objCustomerNumberChangeSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve customerNumberChangeData(1, i - 2)
    customerNumberChangeData(0, i - 2) = objCustomerNumberChangeSheet.cells(i, 2)
    customerNumberChangeData(1, i - 2) = objCustomerNumberChangeSheet.cells(i, 3)
  Next

  'End of data, clear memory
  Set objCustomerNumberChangeSheet = Nothing

  objWorkbook.close
  objExcel.workbooks.close
  objExcel.quit

  Set objWorkbook = Nothing
  Set objExcel = Nothing
  Set objFile = Nothing
  Set objFSO = Nothing
End Function

Sub CustomerChange()
  'Change customer number if it was found
  For i = 0 To UBound(customerNumberChangeData, 2)
    If cLng(customerNumber) = customerNumberChangeData(0, i) Then customerNumber = customerNumberChangeData(1, i)
  Next
End Sub

Sub CustomerPrice()
  'Load customer pricing if available
  For i = 0 To UBound(customerPricingData, 2)
    If cLng(customerNumber) = customerPricingData(0, i) Then
      laborRate = customerPricingData(1, i)
      partsMarkup = customerPricingData(2, i)
      partsMarkupCap = customerPricingData(3, i)
    End If
  Next
End Sub

Sub TruckPrice()
	'Change pricing if applicable
	For i = 0 To UBound(vehiclePricingData,2)
		If cLng(unitNumber) = vehiclePricingData(0, i) Then
			laborRate = vehiclePricingData(1, i)
			partsMarkup = vehiclePricingData(2, i)
			partsMarkupCap = vehiclePricingData(3, i)
		End If
	Next
End Sub

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If


On Error Resume Next
session.findById("wnd[0]").maximize


Dim a, b, c, d, x, i, l, p, an, message, invoice, RepairOrderNumber, laborRate, partsMarkup, unitNumber, VIN, customerNumber, vehicleDistributionChannel, partsMarkupCap, driveTime, test, jobTotal, lastJob, JobN(), Job(), Sto(), LabT, Lab(), LJob(), LabC, LabCP, PJob(), PQty(), PrtN(), PrtD(), PCst(), customerPricingData(), vehiclePricingData(), customerNumberChangeData()

'User input for invoice
Do Until Len(invoice) = 10
If invoice = "" Then
	invoice = InputBox("What is the invoice number from RTC?", "RTC Invoice")
Else
	invoice = InputBox("This was an invalid format for an invoice." & vbCr & "Please try again.", "RTC Invoice", invoice)
End If
If invoice = "" Then
	WScript.Quit
End If
invoice = Trim(invoice)
Loop

lastJob = Int(InputBox("What is the job number of the last job on the invoice?","Last Job"))
If lastJob = "" Then
	WScript.Quit
End If

laborRate = 140
partsMarkup = .65
partsMarkupCap = ""

driveTime = ""
' driveTime = Replace(InputBox("Is there drive time?" + vbCr + "If so, what is the job number?"  + vbCr + "If there are multiple, separate the jobs with a comma. (i.e. 1,3)" + vbCr + "If not, leave blank.","Drive Time")," ","")


Sub CheckCustomer()
	session.findById("wnd[0]").sendVKey 0
	Do Until session.findById("wnd[0]/sbar").text = ""
		session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = InputBox("You've entered an invalid customer number." + vbCr + "Please retry.","Customer Number")
		session.findById("wnd[0]").sendVKey 0
	Loop
End Sub
Sub InDT(x)
	On Error Resume Next
	If IsEmpty(driveTime(0)) Then
		If Int(driveTime) = Int(x) Then
			an = 1
		End IF
	Else
		For z = 0 to UBound(driveTime)
			If driveTime(z) = Int(x) Then
				an = 1
			End If
		Next
	End IF
	Err.Clear
End Sub
Sub GrabLabor()
	For x = 0 To jobTotal - 1
		If Int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"JOBS")) = JobN(x) Then
			LJob(l) = x + 1
		End IF
	Next
	If LJob(l) = "" Then
		LJob(l) = 1
	End If
	Lab(l) = Lab(l) + Round(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZMENG"),2)
	LabT = LabT + Abs(Lab(l))
	l = l + 1
End Sub
Sub GrabPartJob(y)
	For x = 0 To jobTotal - 1
		If Int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"JOBS")) = JobN(x) Then
			PJob(p) = x + 1
		End IF
	Next
	If PJob(p) = "" Then
		PJob(p) = 1
	End If
End Sub
Sub GrabPart()
	PQty(p) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZMENG")
	PrtN(p) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ITOBJID")
	PrtD(p) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"DESCR1")
	PCst(p) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZZXWAVWR")
End Sub



'Verify choises
If MsgBox("Invoice Number: " & invoice & vbCr & "Last Job Number: " & lastJob, vbYesNo, "Verify the following information.") = vbNo Then
	WScript.Quit
End If


If driveTime = "" Then
	driveTime = 0
Else
	If InStr(driveTime,",") > 0 Then
		driveTime = Split(driveTime,",")
		For i = 0 to UBound(driveTime)
			driveTime(i) = Int(driveTime(i))
		Next
	End If
End If

loadConfigFile()


'Go to invoice
session.findById("wnd[0]/tbar[0]/okcd").text = "/NFB03"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/txtRF05L-BELNR").text = invoice
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").setCurrentCell 0,"KOBEZ"
If Err.Number <> 0 Then
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	MsgBox("This is not an invoice.")
	WScript.Quit
End If


'Grab Labor Cost
i = 0
Do Until Err.Number <> 0
	session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").setCurrentCell i,"KOBEZ"
	If inStr(session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getCellValue(i,"KOBEZ"),"Cost") > 0 Then
		If Not inStr(session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getCellValue(i,"ZUONR"),"DBM") > 0 Then
			LabC = LabC + Abs(session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getCellValue(i,"AZBET"))
		Else
			RepairOrderNumber = Right(Left(session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").getCellValue(i,"ZUONR"),13),8)
		End If
	End If
	i = i + 1
Loop
Err.Clear

If RepairOrderNumber = "" Then
	RepairOrderNumber = InputBox("Couldn't find the RO Number." & vbCr & "What is it?")
End If

'Go to RTC RO to fill 
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = RepairOrderNumber
session.findById("wnd[0]").sendVKey 0

unitNumber = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txtIS_VLCACTDATA_ITEM-ZZUN").text
VIN = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txt/DBM/VEHORDCOM-VHVIN").text


'Check if it's a rental
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/btnBUTTON_VEH_SHOW").press
vehicleDistributionChannel = session.findById("wnd[0]/usr/tabsMAIN/tabpVEHDETAIL/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:SAPLZZGC001_01:7100/tabsDATAENTRY/tabpDATAENTRY_FC1/ssubDATAENTRY_SCA:SAPLZZGC001_01:9100/ctxtVLCDIAVEHI-DBM_VTWEG").text
session.findById("wnd[0]/tbar[0]/btn[3]").press


'Read Jobs
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select

jobTotal = 0
a = 0
b = 0
c = 1
d = 0
an = 0
Do Until Int(d & a & b & c) > lastJob
	InDT(Int(d & a & b & c))
	If an = 0 Then
		session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA4:/DBM/SAPLORDER_UI:2053/subSUBSCREEN_2053:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2323/cntlTREE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "J00" & d & a & b & c,"1"
		If Not Err.Number <> 0 Then
			If Not "CORES" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text Then
				If Not "BACK ORDER JOB" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text Then
					Redim Preserve Job(jobTotal), JobN(jobTotal), Sto(jobTotal)
					Job(jobTotal) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text
					JobN(jobTotal) = Int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-JOBNR").text)
					session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/btnJOB_LONG_TEXT").press
					Sto(jobTotal) = session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text
					session.findById("wnd[1]/tbar[0]/btn[12]").press
					jobTotal = jobTotal + 1
				End If
			End If
		End If
	Else
		an = 0
	End If
	Err.Clear
	c = c + 1
	If c = 10 Then
		b = b + 1
		c = 0
	End If
	If b = 10 Then
		a = a + 1
		b = 0
	End If
	If a = 10 Then
		d = d + 1
		a = 0
	End If
Loop


'Read Labor and parts
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
i = 0
l = 0
p = 0
LabT = 0
Err.Clear

Do Until Err.Number <> 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell i,"ITCAT"
If Not session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ITCAT") = "" Then
	If session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ITCAT") = "P001" Then
		ReDim Preserve Lab(l),LJob(l)
		InDT(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"JOBS"))
		If an = 0 Then
			GrabLabor()
		Else
			If l = 0 Then
				Lab(0) = Round(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZMENG"),1)
				LJob(0) = 1
			Else
				Lab(l - 1) = Lab(l - 1) + Round(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZMENG"),1)
				LabT = LabT + Round(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZMENG"),1)
			End If
			an = 0
		End If
		LabC = LabC + Round(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ZZXWAVWR"),2)
	Else
		If Not session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"ITCAT") = "ZCOR" Then
			ReDim Preserve PJob(p),PQty(p),PrtN(p),PrtD(p),PCst(p)
			GrabPart()
			InDT(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(i,"JOBS"))
			If an = 0 Then
				GrabPartJob(p)
			Else
				PJob(p) = 1
				an = 0
		End If
				p = p + 1
		End If
	End If
End If
i = i + 1
Loop
LabCP = LabC / LabT
Err.Clear


'Create a new RO
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-VHVIN").text = VIN
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON05").press
If session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = "100000" Then
	temp = InputBox("What is the customer number?","Customer")
	If temp = "" Then
		WScript.Quit
	End If
	session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = temp
	CheckCustomer()
End If
customerNumber = session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text
CustomerChange()
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = customerNumber
CustomerPrice()
TruckPrice()
If vehicleDistributionChannel = 16 Then
	temp = InputBox("This is a rental." + vbCr + "What is the customer number?","Customer")
	If temp = "" Then
		WScript.Quit
	End If
	session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = temp
	laborRate = 140
	partsMarkup = .65
	partsMarkupCap = ""
	CheckCustomer()
End If
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/cmb/DBM/ORDER_SEARCH-AUFART").key = "ZS00"
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PERNR").text = advisorNumber
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON03").press


'Error Check
message = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0,"T_MSG")
If message <> "" Then
	Do Until message = ""
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = InputBox(message + vbCr + "Please select a different customer number","Customer Number")
		CheckCustomer()
		session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON03").press
		message = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0,"T_MSG")
		If Err.Number <> 0 Then
			message = ""
		End If
	Loop
End If

session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear


'Header
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text

session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "Rebill"
session.findById("wnd[1]/tbar[0]/btn[8]").press

session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA7:/DBM/SAPLORDER_UI:2070/btnGS_ORDER_SCREENS-SCARCP_ICON").press
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA7:/DBM/SAPLORDER_UI:2070/subSUBSCREEN_2070:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:SAPLZZMM031_PARTS:2010/txt/DBM/VBAK_COM-ZZRTC_ORDNO").text = RepairOrderNumber
session.findById("wnd[0]").sendVKey 11
message = session.findById("wnd[1]/usr/txtMESSTXT1").text
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
message = session.findById("wnd[1]/usr/txtMESSTXT1").text
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
If InStr(message,"already used") > 0 Then
	session.findById("wnd[0]/mbar/menu[0]/menu[1]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/btnBUTTON_1").press
	session.findById("wnd[0]/tbar[0]/btn[3]").press
	MsgBox("This rebill was already done." + vbCr + "Reference RO " & Right(message,8))
	WScript.Quit
End If


'Fill Jobs Out
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select

message = session.findById("wnd[1]/usr/txtMESSTXT1").text
session.findById("wnd[1]/tbar[0]/btn[0]").press
Err.Clear
session.findById("wnd[0]").sendVKey 0

For i = 0 To jobTotal - 1
	If Not Job(i) = "" Then
		session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = Job(i)
		session.findById("wnd[0]").sendVKey 0
		If Not Sto(i) = "" Then
			session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/btnJOB_LONG_TEXT").press
			session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = Sto(i)
			session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "0002","COLUMN1"
			session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = Sto(i)
			session.findById("wnd[1]/tbar[0]/btn[8]").press
		End If
		session.findById("wnd[0]").sendVKey 5
	End If
Next

'Fill Out Labor and Parts
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/cmb/DBM/S_POS-ITCAT").key = "ZSUB"

'Labor
For i = 0 To l - 1
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "LABOR REBILL"
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-ZMENG").text = Lab(i)
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-DESCR1").text = "Mobile Labor"
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-KBETM").text = laborRate
	if vehicleDistributionChannel = 18 then
		session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-MATNR18").text = "REBILLSUB-CM"
	else
		session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-MATNR18").text = "REBILLSUB"
	end if
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = LJob(i)
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-REBATE").text = Round(LabCP * Lab(i),2)
	session.findById("wnd[0]").sendVKey 0
Next

'Parts
For i = 0 To p - 1
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PARTS REBILL"
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-ZMENG").text = PQty(i)
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-DESCR1").text = Left(PrtN(i) & "  " & PrtD(i),40)
	If InStr(PrtN(i),"PARTSBUYOUT") + InStr(PrtN(i),"SUBLET") > 0 Then
		session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-KBETM").text = Round(Round(PCst(i) / PQty(i),2) * 1.2,2)
	Else
		If Not partsMarkupCap = "" Then
			If Not Round(Round(PCst(i) / PQty(i),2) * (1 + partsMarkup),2) > Round(PCst(i) / PQty(i),2) + partsMarkupCap Then
				session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-KBETM").text = Round(Round(PCst(i) / PQty(i),2) * (1 + partsMarkup),2)
			Else
				session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-KBETM").text = Round(PCst(i) / PQty(i),2) + partsMarkupCap
			End If
		Else
			session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-KBETM").text = Round(Round(PCst(i) / PQty(i),2) * (1 + partsMarkup),2)
		End If
	End If
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-MATNR18").text = "REBILLSUB"
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = PJob(i)
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-REBATE").text = PCst(i)
	session.findById("wnd[0]").sendVKey 0
Next

'Add labor cost adjustment to first line
If Not LabC - (LabCP * LabT) = 0 Then
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell 0,"ITCAT"
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell
	session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-REBATE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-REBATE").text + LabC - Round(LabCP * LabT,2)
	session.findById("wnd[0]").sendVKey 0
End IF

session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[0]/tbar[1]/btn[37]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press

If Not message = "" Then
	MsgBox(message)
End If