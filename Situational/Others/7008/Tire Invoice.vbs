Dim advisorNumber, branch
advisorNumber = "4608"
branch = "7008"

Dim invoiceNumber, purchaseOrderNumber, orderType, invoiceCost, invoiceHasTires, unitNumber, repairOrderNumber, laborCost, purchaseReq, vendorNumber, jobDescription, roShouldBeClosed, firstOpenPoLine
roShouldBeClosed = true

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

  ' If branch isn't provided, ask for it
  if branch = "" then
    branch = inputBox("What is the branch this is for?", "Branch")
    if branch = "" then
      WScript.Quit
    elseif len(branch) <> 4 or not isNumeric(branch) then
      msgBox "Please enter a valid branch.", 0, "Error"
      branch = ""
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
    jobDescription = inputBox("What would you like to be the job name?", "Job Name")
    if jobDescription = "" then
      WScript.Quit
    end if
  end if

  ' Get the invoice number
  if invoiceNumber = "" then
    invoiceNumber = inputBox("What is the invoice number?", "Invoice Number")
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
  
  ' If it's retail, ask if there are tires
  if invoiceHasTires = "" then
    invoiceHasTires = false
    if orderType = "Retail" then
      if msgBox("Are there tires on this invoice?", vbYesNo, "Tires") = 6 then
        invoiceHasTires = true
      else
        laborCost = invoiceCost
      end if
    end if
  end if
  
  ' If the invoice has tires on it, we need to know the cost for just labor
  if invoiceHasTires and laborCost = "" then
    laborCost = inputBox("What is the cost for labor? (Invoice total minus the cost of tires)", "Labor Cost")
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

  if not validate() then
    exit function
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
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").select
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").select
  headerText = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text
  headerText = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text

  conditions = split(split(headerText, vbCr)(0),"|")
  vendorNumber = conditions(0)
  unitNumber = replace(conditions(1),"-","")
  orderType = conditions(2)
end function

function roTypeIsVerrified()
  roTypeIsVerrified = false
  if orderType = "Other" then
    orderType = inputBox("What RO type would you like to open this under?" & vbCr & "1) Internal" & vbCr & "2) Retail" & vbCr & "3) VIO", "RO Type")
    if orderType = "1" then
      orderType = "Internal"
    elseif orderType = "2" then
      orderType = "Retail"
    elseif orderType = "3" then
      orderType = "VIO"
    else
      orderType = "Other"
      msgBox "Please enter a valid option.", 0, "Error"
      exit function
    end if
  end if
  roTypeIsVerrified = true
end function

function validate()
  validateMessage = "Is all of the following information correct?" & vbCr & vbCr & "Advisor Number:" & vbCr & advisorNumber & vbCr & vbCr & "Job Title:" & vbCr & jobDescription & vbCr & vbCr & "Unit Number:" & vbCr & unitNumber & vbCr & vbCr & "RO Type:" & vbCr & orderType & vbCr & vbCr & "Vendor Number:" & vbCr & vendorNumber & vbCr & vbCr & "Invoice Number:" & vbCr & invoiceNumber & vbCr & vbCr & "Total Invoice Amount:" & vbCr & invoiceCost
  if orderType = "Retail" and invoiceHasTires then
    validateMessage = validateMessage & vbCr & vbCr & "Labor Cost:" & vbCr & laborCost & vbCr & vbCr & "Tire Cost:" & vbCr & (invoiceCost - laborCost)
  end if

  if msgBox(validateMessage, vbYesNo, "Validate") = vbNo then
    validateMessage = "Which entry would you like to change?" & vbCr & vbCr & "1) Advisor Number:" & vbCr & advisorNumber & vbCr & vbCr & "2) Job Title:" & vbCr & jobDescription & vbCr & vbCr & "3) Unit Number:" & vbCr & unitNumber & vbCr & vbCr & "4) RO Type:" & vbCr & orderType & vbCr & vbCr & "5) Vendor Number:" & vbCr & vendorNumber & vbCr & vbCr & "6) Invoice Number:" & vbCr & invoiceNumber & vbCr & vbCr & "7) Total Invoice Amount:" & vbCr & invoiceCost
    if orderType = "Retail" then
      validateMessage = validateMessage & vbCr & vbCr & "8) Invoice Has Tires:" & vbCr & invoiceHasTires
      if invoiceHasTires then
        Dim changeEntryOption
        validateMessage = validateMessage & vbCr & vbCr & "9) Labor Cost:" & vbCr & laborCost
      end if
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
      vendorNumber = inputBox("What vendor is this for?" & vbCr & "1) Michelin" & vbCr & "2) Bridgestone / Yokohama / Continental" & vbCr & "3) Southern Tire Mart (labor)" & vbCr & "4) Jack's Tires (labor)" & vbCr & "5) Border Tire (labor)", "Vendor")
      if vendorNumber = "" then
        WScript.Quit
      elseif vendorNumber = "1" then
        vendorNumber = "214567"
      elseif vendorNumber = "2" then
        vendorNumber = "200524"
      elseif vendorNumber = "3" then
        vendorNumber = "207488"
      elseif vendorNumber = "4" then
        vendorNumber = "209330"
      elseif vendorNumber = "5" then
        vendorNumber = "242462"
      else
        vendorNumber - ""
        msgBox "Please enter a valid option.", 0, "Error"
        exit function
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
  on error resume next
  ' Select the order tab and click on appropriate order
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
  if orderType = "Internal" then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
  elseif orderType = "Retail" then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS00","Column01"
  elseif orderType = "VIO" then
    session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS15","Column01"
  end if
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

function makeRepairOrder()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unitNumber
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press

  if session.findById("wnd[0]/sbar").text = "No vehicles could be selected" then
    unitNumber = inputBox("You entered an invalid unit Number." & vbCr & "What is the correct unit?")
    makeRO = false
    exit function
  else
    do while vehicleIsLocked()
    loop
  end if

  ' Header
  ' Set mileage and hours to previous numbers
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
  ' VIO needs the account assignment category
  if orderType = "VIO" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "901"
    if session.findById("wnd[0]/sbar").text = "AAC 901 not allowed for Vehicle Status P500" then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "902"
      msgBox "You'll need to come back later to change the Account Assignment Category to 901 after the vehicle is in P200 status.", 0, "VIO"
      roShouldBeClosed = false
    end if
  end if

  ' Go to the job tab and fill out job 1 as tires
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = jobDescription
  session.findById("wnd[0]").sendVKey 0

  ' Go to the item tab and fill out the purchase req(s)
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/cmb/DBM/S_POS-ITCAT").key = "P010"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETNT"
  session.findById("wnd[0]").sendVKey 0
  if orderType = "Internal" or orderType = "VIO" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "TIRES"
  elseif orderType = "Retail" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "LALBOR"
  end if
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-JOBS").text = "1"
  if orderType = "Retail" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-KBETM").text = round(laborCost * 1.15, 2)
  end if
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-VERPR").text = laborCost
  session.findById("wnd[0]").sendVKey 0
  if invoiceHasTires and orderType = "Retail" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-ITOBJID").text = "SUBLETTI"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-DESCR1").text = "TIRES"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/ctxt/DBM/S_POS-JOBS").text = "1"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-VERPR").text = (invoiceCost - laborCost)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3319/txt/DBM/S_POS-KBETM").text = round((invoiceCost - laborCost) * 1.15, 2)
    session.findById("wnd[0]").sendVKey 0
  end if

  ' Go to the parts tab and enter the vendor number
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell 0,"LIFNR",vendorNumber
  if invoiceHasTires and orderType = "Retail" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell 1,"LIFNR",vendorNumber
  end if
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressEnter
  session.findById("wnd[0]").sendVKey 11
  if orderType = "VIO" then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
  end if
  
  ' Create the purchase req
  if invoiceHasTires and orderType = "Retail" then
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectedRows = "0-1"
  else
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectedRows = "0"
  end if
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  if orderType = "Internal" or orderType = "VIO" then
    session.findById("wnd[1]/usr/lbl[23,17]").setFocus
  elseif orderType = "Retail" then
    session.findById("wnd[1]/usr/lbl[23,18]").setFocus
  end if
  session.findById("wnd[1]").sendVKey 2
  if orderType = "VIO" then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
  end if
  session.findById("wnd[0]/tbar[1]/btn[13]").press

  ' Store req and ro number
  purchaseReq = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(0,"ZZBANFN")
  if purchaseReq = "" then
    msgBox "Something went wrong, please finish this manually"
    Wscript.Quit
  end if
  if orderType = "Internal" or orderType = "VIO" then
    repairOrderNumber = right(left(replace(session.findById("wnd[0]/titl").text, "&", ""), 37), 8)
  elseif orderType = "Retail" then
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
  if invoiceHasTires and orderType = "Retail" then
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[27," & firstOpenPoLine & "]").text = "10"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BANFN[13," & firstOpenPoLine + 1 & "]").text = purchaseReq
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-BNFPO[27," & firstOpenPoLine + 1 & "]").text = "20"
  end if
  session.findById("wnd[0]").sendVKey 0

  ' Select first line and delete
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").getAbsoluteRow(0).selected = true
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/btnDELETE").press
  session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

  ' Wait for the user to finish entering the tires
  if invoiceHasTires and orderType = "Retail" then
  elseif invoiceHasTires then
    msgBox "DO NOT PRESS OK!" & vbCr & "Please add all tires to the PO." & vbCr & "AFTER you do that, press the OK button.", 0, "WAIT!!!!"
  end if

  ' Verify that the invoice total matches the PO total
  do until invoiceTotalCostMatchesPurchaseOrderTotal()
    msgBox "The PO doesn't match the invoice total." & vbCr & "Please verify your pricing on the PO and press OK to try again." & vbCr & "The invoice total is: " & invoiceCost
  loop

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

  ' If it's internal, MIGO and close
  ' MIGO
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell 0,"ZZMIGO"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressButtonCurrentCell
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").selected = true
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").setFocus
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/subSUB_BUTTONS:SAPLMIGO:0210/btnOK_TAKE_VALUE").press
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = invoiceNumber
  session.findById("wnd[0]/tbar[1]/btn[23]").press

  if orderType = "Internal" or orderType = "VIO" and roShouldBeClosed then
    ' Open RO, release, create billing
    session.findById("wnd[0]/tbar[1]/btn[13]").press
    session.findById("wnd[0]/tbar[1]/btn[37]").press
    session.findById("wnd[0]/tbar[1]/btn[40]").press
  end if
  if orderType = "VIO" and roShouldBeClosed then
    session.findById("wnd[1]/tbar[0]/btn[0]").press
  end if

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