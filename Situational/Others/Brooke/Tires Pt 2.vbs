dim advisorNumber, branch
advisorNumber = "2676"

dim invoiceNumber, purchaseOrderNumber, orderType, invoiceCost, invoiceHasTires, unitNumber, repairOrderNumber, laborCost, purchaseReq, vendorNumber, jobDescription, jobName, roShouldBeClosed, firstOpenPoLine, numberOfLines, markup, isRental
dim vendorOptions()
roShouldBeClosed = true

function loadConfigFile()
  'Find the path name of this script
  strPath = Wscript.scriptFullname
  'Create File System Object
  set objFSO = createObject("Scripting.FileSystemObject")
  'Create object for this script's file
  set objFile = objFSO.getFile(strPath)
  'Get the folder that this script is located in
  strFolder = objFSO.getParentFolderName(objFile)
  'Create an Excel Object
  set objExcel = createObject("Excel.Application")
  'Open the Rebill Pricing Excel File
  set objWorkbook = objExcel.workBooks.open(strFolder & "\Tire Vendors.xlsx")

  dim rowcount
  '-----------------------------Vendors-----------------------------
  'Load the sheet and store the data
  set vendorsSheet = objWorkbook.worksheets("Vendors")
  rowcount = vendorsSheet.Usedrange.Rows.Count

  for i = 2 to rowcount
    redim preserve vendorOptions(1, i - 2)
    vendorOptions(0, i - 2) = vendorsSheet.cells(i, 1)
    vendorOptions(1, i - 2) = vendorsSheet.cells(i, 2)
  next

  'End of data, clear memory
  set vendorsSheet = Nothing

  objWorkbook.close
  objExcel.workbooks.close
  objExcel.quit

  set objWorkbook = Nothing
  set objExcel = Nothing
  set objFile = Nothing
  set objFSO = Nothing
end function

function isValidCostFormat(cost)
  if not isNumeric(cost) then
    isValidCostFormat = false
    exit function
  end if
  if inStr(cost, ".") then
    if len(split(cost, ".")(1)) > 2 then
      isValidCostFormat = false
      exit function
    end if
  end if
  isValidCostFormat = true
end function

function askForUserInput()
  askForUserInput = false
  
  ' If Advisor Number isn't provided, ask for it
  if advisorNumber = "" then
    advisorNumber = inputBox("What is your advisor number?", "Advisor Number")
    if advisorNumber = "" then
      WScript.Quit
    elseif not isNumeric(advisorNumber) then
      msgBox "Please enter a number.", 0, "Error"
      advisorNumber = ""
      exit function
    end if
  end if

  ' Get the PO number
  if purchaseOrderNumber = "" then
    purchaseOrderNumber = inputBox("What is the purchase order number?", "PO Number")
    if purchaseOrderNumber = "" then
      WScript.Quit
    elseif len(purchaseOrderNumber) <> 10 then
      msgBox "Please enter a valid PO number.", 0, "Error"
      purchaseOrderNumber = ""
      exit function
    end if
  end if

  goToPOForConfigInformation()

  do until roTypeIsVerrified()
  loop

  ' Get the job title
  if jobDescription = "" then
    jobDescription = inputBox("What would you like to be the job name?", "Job Name", jobName)
    if jobDescription = "" then
      WScript.Quit
    end if
  end if

  ' Get the invoice number
  if invoiceNumber = "" then
    invoiceNumber = uCase(inputBox("What is the invoice number?", "Invoice Number"))
    if invoiceNumber = "" then
      WScript.Quit
    end if
  end if

  ' Get the invoice total
  if invoiceCost = "" then
    invoiceCost = inputBox("What is the toal cost of the invoice?", "Invoice Total")
    invoiceCost = replace(invoiceCost, "$", "")
    if invoiceCost = "" then
      WScript.Quit
    else
      if not isValidCostFormat(invoiceCost) then
        msgBox "Please give a valid cost format.", 0, "Error"
        invoiceCost = ""
        exit function
      end if
    end if
    invoiceCost = cDbl(invoiceCost)
  end if
  
  ' If it's not a VIO, ask if there are tires
  if not isVIO() and invoiceHasTires = "" then
    if msgBox("Are there tires on this invoice?", vbYesNo, "Tires") = 6 then
      invoiceHasTires = true
    else
      invoiceHasTires = false
      laborCost = invoiceCost
    end if
  else
    if laborCost = "" then
      laborCost = invoiceCost
    end if
  end if
  
  ' If the invoice has tires on it, we need to know the cost for just labor
  if laborCost = "" then
    laborCost = inputBox("What is the cost for labor? (Invoice total minus the cost of all tires)", "Labor Cost")
    if laborCost = "" then
      WScript.Quit
    end if
    do until isvalidCostFormat(laborCost)
      laborCost = inputBox("You've entered an invalid currency format. Please enter a valid cost.", "Invalid Cost", laborCost)
      if laborCost = "" then
        WScript.Quit
      end if
    loop
  end if

  if laborCost = "" then
    laborCost = invoiceCost
  end if

  if isRetail() and markup = "" then
    markup = inputBox("What would you like the markup to be?" & vbCr & "Just put the number (i.e. 20)", "Markup")
    if markup = "" then
      wScript.quit
    end if
    if not isNumeric(markup) then
      msgBox "Please enter a number.",, "Invalid"
      markup = ""
      exit function
    end if
  end if

  if not validate() then
    exit function
  end if

  if markup <> "" then
    markup = cDbl("1." & markup)
  end if
  
  ' If all input is received, return true to move on
  askForUserInput = true
end function

function goToPOForConfigInformation()
  on error resume next
  if vendorNumber <> "" then
    exit function
  end if

  ' Go to the PO
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/tbar[1]/btn[17]").press
  session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = purchaseOrderNumber
  session.findById("wnd[1]").sendVKey 0

  ' Grab header text to extract configs
  itteration = findItteration()
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").select
  headerText = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text
  branch = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[10,0]").text

  conditions = split(split(headerText, vbCr)(0),"|")
  vendorNumber = conditions(0)
  unitNumber = replace(conditions(1),"-","")
  orderType = conditions(2)
  jobName = trim(split(headerText, vbCr)(1))
end function

function roTypeIsVerrified()
  roTypeIsVerrified = false
  if orderType = "Other" then
    orderType = inputBox("What RO type would you like to open this under?" & vbCr & "1) Internal" & vbCr & "2) Retail" & vbCr & "3) VIO", "RO Type")
    if orderType = "1" then
      isInternal()
    elseif orderType = "2" then
      isRetail()
    elseif orderType = "3" then
      isVIO()
    else
      orderType = "Other"
      msgBox "Please enter a valid option.", 0, "Error"
      exit function
    end if
  end if
  roTypeIsVerrified = true
end function

function validate()
  validateMessage = "Is all of the following information correct?" & vbCr & vbCr & "Advisor Number:" & vbCr & advisorNumber & vbCr & vbCr & "Job Title:" & vbCr & jobDescription & vbCr & vbCr & "Unit Number:" & vbCr & unitNumber & vbCr & vbCr & "RO Type:" & vbCr & orderType & vbCr & vbCr & "Vendor Number:" & vbCr & vendorNumber & vbCr & vbCr & "Invoice Number:" & vbCr & invoiceNumber & vbCr & vbCr & "Total Invoice Amount:" & vbCr & "$" & invoiceCost
  if invoiceHasTires then
    validateMessage = validateMessage & vbCr & vbCr & "Labor Cost:" & vbCr & "$" & laborCost & vbCr & vbCr & "Tire Cost:" & vbCr & "$" & (invoiceCost - laborCost)
  end if
  if isRetail() then
    validateMessage = validateMessage & vbCr & vbCr & "Markup:" & vbCr & markup & "%"
  end if
  
  if msgBox(validateMessage, vbYesNo, "Validate") = vbNo then
    Dim changeEntryOption
    validateMessage = "Which entry would you like to change?" & vbCr & vbCr & "1) Advisor Number:" & vbCr & advisorNumber & vbCr & vbCr & "2) Job Title:" & vbCr & jobDescription & vbCr & vbCr & "3) Unit Number:" & vbCr & unitNumber & vbCr & vbCr & "4) RO Type:" & vbCr & orderType & vbCr & vbCr & "5) Vendor Number:" & vbCr & vendorNumber & vbCr & vbCr & "6) Invoice Number:" & vbCr & invoiceNumber & vbCr & vbCr & "7) Total Invoice Amount:" & vbCr & "$" & invoiceCost
    if invoiceHasTires then
      validateMessage = validateMessage & vbCr & vbCr & "8) Invoice Has Tires:" & vbCr & invoiceHasTires & vbCr & vbCr & "9) Labor Cost:" & vbCr & "$" & laborCost
    end if
    if isRetail() then
      validateMessage = validateMessage & vbCr & vbCr & "10) Markup:" & vbCr & markup & "%"
    end if
    changeEntryOption = inputBox(validateMessage,"Validate")
    if changeEntryOption = "1" then
      ' Advisor Number
      advisorNumber = ""
    elseif changeEntryOption = "2" then
      ' Job Title
      jobDescription = ""
    elseif changeEntryOption = "3" then
      ' Unit Number
      unitNumber = replace(inputBox("What would you like to change the unit number to?","Unit Number Change"), "-", "")
    elseif changeEntryOption = "4" then
      ' RO Type
      orderType = inputBox("What would you like to change the RO type to?" & vbCr & "1) Internal" & vbCr & "2) Retail" & vbCr & "3) VIO")
      if orderType = "1" then
        orderType = "Internal"
      elseif orderType = "2" then
        orderType = "Retail"
      elseif orderType = "3" then
        orderType = "VIO"
      else
        orderType = "Other"
        msgBox "Please enter a valid option.", 0, "Error"
      end if
    elseif changeEntryOption = "5" then
      ' Vendor Number
      loadConfigFile()
      for i = 0 to uBound(vendorOptions, 2)
        vendorMessage = vendorMessage & i + 1 & ") " & vendorOptions(0, i) & vbCr
      next
      vendorNumber = inputBox("What vendor is this for?" & vbCr & vendorMessage, "Vendor Number")
      if vendorNumber = "" then
        WScript.Quit
      else
        if isNumeric(vendorNumber) then
          vendorNumber = cInt(vendorNumber)
        end if
        if isNumeric(vendorNumber) and vendorNumber <= uBound(vendorOptions, 2) + 1 and vendorNumber > 0 then
          vendorNumber = vendorOptions(1, vendorNumber - 1)
        else
          vendorNumber = ""
          msgBox "Please enter a valid option.", 0, "Error"
        end if
      end if
    elseif changeEntryOption = "6" then
      ' Invoice Number
      invoiceNumber = ""
    elseif changeEntryOption = "7" then
      ' Invoice Cost
      invoiceCost = ""
    elseif changeEntryOption = "8" then
      ' Invoice Has Tires
      invoiceHasTires = ""
      laborCost = ""
    elseif changeEntryOption = "9" then
      ' Labor Cost
      laborCost = ""
    elseif changeEntryOption = "10" then
      ' Markup
      markup = ""
    elseif changeEntryOption = "" then
      WScript.Quit
    else
      ' Invalid Entry
    end if
    validate = false
    exit function
  end if
  validate = true
end function

function vehicleIsLocked()
  dim customerNumber
  on error resume next
  ' Select the order tab and click on appropriate order
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
  if isInternal() then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
  elseif isRetail() then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS00","Column01"
  elseif isVIO() then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS15","Column01"
  end if
  customerNumber = session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text
  if isRetail() and isRental then
    customerNumber = inputBox("This is a rental. What is the customer number that you want to bill?", "Customer Number", customerNumber)
  end if
  if isRetail() and customerNumber = "100000" then
    customerNumber = inputBox("This is set to Rush Internal. What is the customer number that you want to bill?", "Customer Number")
  end if
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = customerNumber
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = advisorNumber
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = branch
  session.findById("wnd[0]").sendVKey 0
  message = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0, "T_MSG")
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  if message <> "" then
    if msgBox(message & vbCr & "Would you like to try again?" & vbCr & "Pressing no cancels this entire process.", vbYesNo, "Try Again") = 6 then
      vehicleIsLocked = true
      exit function
    else
      Wscript.Quit
    end if
  end if
  vehicleIsLocked = false
end function

function isInternal()
  isInternal = (orderType = "Internal")
end function

function isRetail()
  isRetail = (orderType = "Retail")
end function

function isVIO()
  isVIO = (orderType = "VIO")
end function

function makeRepairOrder()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  ' Unit Number
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unitNumber
  ' VIN
  ' session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtVHVIN-LOW").text = unitNumber
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press

  if session.findById("wnd[0]/sbar").text = "No vehicles could be selected" then
    unitNumber = inputBox("You entered an invalid unit Number." & vbCr & "What is the correct unit?")
    makeRO = false
    exit function
  else
    if "16" = session.findById("wnd[0]/usr/tabsMAIN/tabpVEHDETAIL/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:SAPLZZGC001_01:7100/tabsDATAENTRY/tabpDATAENTRY_FC1/ssubDATAENTRY_SCA:SAPLZZGC001_01:9100/ctxtVLCDIAVEHI-DBM_VTWEG").text then
      isRental = true
    end if
    do while vehicleIsLocked()
    loop
  end if

  ' Header
  ' Set mileage and hours to previous numbers
  if session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text = "0" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = "1"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = "1"
  else
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
  end if
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
  ' VIO needs the account assignment category
  if isVIO() then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "901"
    if session.findById("wnd[0]/sbar").text = "AAC 901 not allowed for Vehicle Status P500" then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "902"
      msgBox "You'll need to come back later to change the Account Assignment Category to 901 after the vehicle is in P200 status.", 0, "VIO"
      roShouldBeClosed = false
    end if
  end if

  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
  session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "Tires"
  session.findById("wnd[1]/tbar[0]/btn[8]").press

  ' Go to the job tab and fill out job 1 as tires
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = jobDescription
  session.findById("wnd[0]").sendVKey 0

  ' Go to the item tab and fill out the purchase req(s)
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/cmb/DBM/S_POS-ITCAT").key = "P010"
  ' ----Labor----
  numberOfLines = 0
  if laborCost <> "0" then
    numberOfLines = numberOfLines + 1
    if isInternal() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETNT"
    end if
    if isRetail() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETTX"
    end if
    if isVIO() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETVI"
    end if
    session.findById("wnd[0]").sendVKey 0
    if isInternal() or isRetail() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "LABOR"
    end if
    if isVIO() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "TIRES"
    end if
    ' Qty
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-JOBS").text = "1"
    if isRetail() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-KBETM").text = round(laborCost * markup, 2)
    end if
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-VERPR").text = laborCost
    session.findById("wnd[0]").sendVKey 0
  end if
  ' ----Tires----
  if invoiceHasTires and not isVIO() then
    numberOfLines = numberOfLines + 1
    if isInternal() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETTI"
    end if
    if isRetail() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETTX"
    end if
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "TIRES"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-JOBS").text = "1"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-VERPR").text = (invoiceCost - laborCost)
    if isRetail() then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-KBETM").text = round((invoiceCost - laborCost) * markup, 2)
    end if
    session.findById("wnd[0]").sendVKey 0
  end if

  ' Go to the parts tab and enter the vendor number
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell 0,"LIFNR",vendorNumber
  if numberOfLines = 2 then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell 1,"LIFNR",vendorNumber
  end if
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressEnter
  session.findById("wnd[0]").sendVKey 11
  if isVIO() then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
  end if
  
  ' Create the purchase req
  if numberOfLines = 2 then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectedRows = "0-1"
  else
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectedRows = "0"
  end if
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  if isInternal() or isVIO() then
    session.findById("wnd[1]/usr/lbl[23,18]").setFocus
  elseif isRetail() then
    session.findById("wnd[1]/usr/lbl[23,19]").setFocus
  end if
  session.findById("wnd[1]").sendVKey 2
  if isVIO() then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
  end if
  session.findById("wnd[0]/tbar[1]/btn[13]").press

  ' Store req and ro number
  purchaseReq = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(0,"ZZBANFN")
  if purchaseReq = "" then
    msgBox "Something went wrong, please finish this manually"
    Wscript.Quit
  end if
  if isInternal() or isVIO() then
    repairOrderNumber = right(left(replace(session.findById("wnd[0]/titl").text, "&", ""), 37), 8)
  elseif isRetail() then
    repairOrderNumber = right(left(replace(session.findById("wnd[0]/titl").text, "&", ""), 29), 8)
  end if

  if len(replace(repairOrderNumber, " ", "")) <> 8 or not isNumeric(repairOrderNumber) then
    repairOrderNumber = inputBox("There was an issue reading the RO number. Please enter it here.")
  end if

  makeRepairOrder = true
end function

function invoiceTotalCostMatchesPurchaseOrderTotal()
  on error resume next
  dim purchaseOrderTotalCost
  purchaseOrderTotalCost = replace(replace(session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/ssubHEADER_CUM_2:SAPLMEGUI:1235/txtMEPO1235-VALUE01").text,",","")," ","")
  purchaseOrderTotalCost = replace(replace(session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1232/ssubHEADER_CUM_2:SAPLMEGUI:1235/txtMEPO1235-VALUE01").text,",","")," ","")
  if cDbl(purchaseOrderTotalCost) = cDbl(invoiceCost) then
    invoiceTotalCostMatchesPurchaseOrderTotal = true
  else
    invoiceTotalCostMatchesPurchaseOrderTotal = false
  end if
end function

function findItteration()
   on error resume next
   dim testCase, a, b
   a = 0
   b = 0
   do while true
      testCase = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & b & a & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,0]").text
      if err.number = 0 then
        exit do
      end if
      err.clear
      a = a + 1
      if a = 10 then
         a = 0
         b = b + 1
      end if
   loop
   findItteration = b & a
end function

function addPurchaseReqToPurchaseOrder()
  on error resume next
  ' Go to the PO
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/tbar[1]/btn[17]").press
  session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = purchaseOrderNumber
  session.findById("wnd[1]").sendVKey 0

  
  ' Open and enter purchase req (this is labor)
  session.findById("wnd[0]/tbar[1]/btn[7]").press
  itteration = findItteration()
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT10").select

  firstOpenPoLine = 0
  do while true
    if session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1," & firstOpenPoLine & "]").text = "" then
      exit do
    end if
    firstOpenPoLine = firstOpenPoLine + 1
  loop

  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[13," & firstOpenPoLine & "]").text = purchaseReq
  if numberOfLines = 2 then
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[27," & firstOpenPoLine & "]").text = "10"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[13," & firstOpenPoLine + 1 & "]").text = purchaseReq
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[27," & firstOpenPoLine + 1 & "]").text = "20"
  end if
  session.findById("wnd[0]").sendVKey 0

  ' Select first line and delete
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").getAbsoluteRow(0).selected = true
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/btnDELETE").press
  session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

  ' Save
  session.findById("wnd[0]/tbar[0]/btn[11]").press
end function

function goToRepairOrderAndMIGO()
  ' Go back to the RO
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = repairOrderNumber
  session.findById("wnd[0]").sendVKey 0

  ' Go to the job tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select

  ' MIGO
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell 0,"ZZMIGO"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressButtonCurrentCell
  on error resume next
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/btnBUTTON_ITEMDETAIL").press
  err.clear
  on error goto 0
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").selected = true
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").setFocus
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/subSUB_BUTTONS:SAPLMIGO:0210/btnOK_TAKE_VALUE").press
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = invoiceNumber
  session.findById("wnd[0]/tbar[1]/btn[23]").press
end function


if Not IsObject(application) Then
  Set SapGuiAuto  = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
End if
if Not IsObject(connection) Then
  Set connection = application.Children(0)
End if
if Not IsObject(session) Then
  Set session    = connection.Children(0)
End if
if IsObject(WScript) Then
  WScript.ConnectObject session,     "on"
  WScript.ConnectObject application, "on"
End if
session.findById("wnd[0]").maximize

do until askForUserInput()
loop
do until makeRepairOrder()
loop
addPurchaseReqToPurchaseOrder()
goToRepairOrderAndMIGO()