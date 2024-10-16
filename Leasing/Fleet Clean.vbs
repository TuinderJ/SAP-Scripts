dim daycabPrice, sleeperPrice, boxTruckPrice, flatBedPrice, feesPrice, truckCount, invoiceNumber, invoiceCost
dim trucks()
' daycabPrice = 43
' sleeperPrice = 48
' boxTruckPrice = 40
' flatBedPrice = 40

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
session.findById("wnd[0]").maximize

' Ask for pricing on daycab, sleeper, box truck
' Ask for unit numbers and types
' Ask for fees (fuel surcharge)
askForUserInput()
createPO()
getPONumberAndMIGO()

' Create PO, vendor and whatnot

' Loop 0 to unit count F|Truck Wash Unit|Type Cost|1111
' Cost center stuff

' unit count + 1 K|Fees|Cost|623000

' Save, Display PO

function askForUserInput()
  if daycabPrice = "" then
    daycabPrice = inputBox("What is the price for Day Cabs?" & vbCr & "This includes the main wash and the frame.", "Day Cab", "43")
  end if
  if sleeperPrice = "" then
    sleeperPrice = inputBox("What is the price for Sleepers?" & vbCr & "This includes the main wash and the frame.", "Sleeper", "48")
  end if
  if boxTruckPrice = "" then
    boxTruckPrice = inputBox("What is the price for Box Trucks?" & vbCr & "This includes the main wash and the frame.", "Box Truck", "40")
  end if
  if flatBedPrice = "" then
    flatBedPrice = inputBox("What is the price for Flat Beds?" & vbCr & "This includes the main wash and the frame.", "Flat Bed", "40")
  end if
  if feesPrice = "" then
    feesPrice = inputBox("What is the price for the fees?" & vbCr & "(Everything else)", "Fees")
  end if
  if invoiceNumber = "" then
    invoiceNumber = inputBox("What is the invoice number?", "Invoice Number")
  end if

  truckCount = -1
  do while true
    dim question, answer
    
    question = "What is the next unit number?" & vbCr & "If you're done, leave it blank."

    if truckCount > -1 then
      for i = 0 to truckCount
        question = question & vbCr & trucks(0, i)
      next
    end if
    answer = inputBox(question, "Unit Number")
    
    if answer = "" then
      exit do
    end if
    truckCount = truckCount + 1
    redim preserve trucks(1, truckCount)
    trucks(0, truckCount) = answer
    
    trucks(1, truckCount) = inputBox("What type of truck is " & answer & "?" & vbCr & _
      "1) Day Cab" & vbCr & _
      "2) Sleeper" & vbCr & _
      "3) Box Truck" & vbCr & _
      "4) Flat Bed" & vbCr & _
      "5) Other" _
      , "Type")
    select case trucks(1, truckCount)
      case "1"
        trucks(1, truckCount) = "Day Cab"
      case "2"
        trucks(1, truckCount) = "Sleeper"
      case "3"
        trucks(1, truckCount) = "Box Truck"
      case "4"
        trucks(1, truckCount) = "Flat Bed"
      case "5"
        trucks(1, truckCount) = "Other"
    end select
  loop

  if truckCount < 0 then
    WScript.Quit
  end if

  invoiceCost = feesPrice
  for i = 0 to truckCount
    select case trucks(1, i)
      case "Day Cab"
        invoiceCost = cDbl(invoiceCost) + cDbl(daycabPrice)
      case "Sleeper"
        invoiceCost = cDbl(invoiceCost) + cDbl(sleeperPrice)
      case "Box Truck"
        invoiceCost = cDbl(invoiceCost) + cDbl(boxTruckPrice)
      case "Flat Bed"
        invoiceCost = cDbl(invoiceCost) + cDbl(flatBedPrice)
      case "Other"
        invoiceCost = cDbl(invoiceCost) + cDbl(inputBox("What is the cost for " & trucks(0, i) & "?", "Misc Price"))
    end select
  next

  dim verifyMessage
  verifyMessage = "Is all of this information correct?" & vbCr & vbCr & _
    "Day Cab Pricing: $" & daycabPrice & vbCr & _
    "Sleeper Pricing: $" & sleeperPrice & vbCr & _
    "Box Truck Pricing: $" & boxTruckPrice & vbCr & _
    "Flat Bed Pricing: $" & flatBedPrice & vbCr & _
    "Fees: $" & feesPrice & vbCr & _
    "Invoice Total: $" & invoiceCost & vbCr & vbCr & _
    "Trucks" & vbCr

  for i = 0 to truckCount
    verifyMessage = verifyMessage + trucks(0, i) & " " & trucks(1, i) & vbCr 
  next
  
  if msgBox(verifyMessage, vbYesNo, "Verify") = vbNo then
    WScript.Quit
  end if

end function

function createPO()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME21N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  
  dim itteration
  itteration = findItteration()
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "243801"
  session.findById("wnd[0]").sendVKey 0
  
  for i = 0 to truckCount
    itteration = findItteration()
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2," & i & "]").text = "F"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[4," & i & "]").text = "Truck Wash " & trucks(0, i)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[5," & i & "]").text = "1"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[6," & i & "]").text = "EA"
    select case trucks(1, i)
      case "Day Cab"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = daycabPrice
      case "Sleeper"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = sleeperPrice
      case "Box Truck"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = boxTruckPrice
      case "Flat Bed"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = flatBedPrice
      case "Other"
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = inputBox("What is the cost for unit " & trucks(0, i) & "?", "Other Price")
    end select
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[9," & i & "]").text = "1111"
  next
  dim lastLine
  lastLine = truckCount + 1
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2," & lastLine & "]").text = "K"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[4," & lastLine & "]").text = "Fees"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[5," & lastLine & "]").text = "1"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[6," & lastLine & "]").text = "EA"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & lastLine & "]").text = feesPrice
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[9," & lastLine & "]").text = "1111"
  session.findById("wnd[0]").sendVKey 0

  dim a, b, c, d
  a = 0
  b = " "
  c = " "
  d = " "
  for i = 0 to truckCount + 1
    a = a + 1
    if a = 10 then
      if b = " " then
        b = 1
      else
        b = b + 1
        if b = 10 then
          if c = " " then
            c = 1
          else
            c = c + 1
            if c = 10 then
              if d = " " then
                c = 1
              else
                c = c + 1
              end if
            end if
          end if
        end if
      end if
    end if

    dim branch
    branch = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & findItteration() & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[10,0]").text
    itteration = findItteration()
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST").key = d & c & b & a
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = "623000"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = branch & "00"

    if i <= truckCount then
      session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR").setFocus
      session.findById("wnd[0]").sendVKey 4
      session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB015/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[0,24]").text = trucks(0, i)
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
    end if
    session.findById("wnd[0]").sendVKey 0
  next

  session.findById("wnd[0]/tbar[0]/btn[11]").press
  session.findById("wnd[0]").sendVKey 3
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

function getPONumberAndMIGO()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME29N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  dim purchaseOrderNumber
  purchaseOrderNumber = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & findItteration() & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").text
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT12/ssubTABSTRIPCONTROL2SUB:SAPLMERELVI:1100/cntlRELEASE_INFO/shellcont/shell").currentCellColumn = "FUNCTION"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT12/ssubTABSTRIPCONTROL2SUB:SAPLMERELVI:1100/cntlRELEASE_INFO/shellcont/shell").clickCurrentCell
  session.findById("wnd[0]/tbar[0]/btn[11]").press

  session.findById("wnd[0]/tbar[0]/okcd").text = "/NMIGO"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = purchaseOrderNumber
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART").text = "101"
  session.findById("wnd[0]").sendVKey 0

  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").selected = true
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").setFocus
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/subSUB_BUTTONS:SAPLMIGO:0210/btnOK_TAKE_VALUE").press
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = invoiceNumber
  session.findById("wnd[0]/tbar[1]/btn[23]").press

  session.findById("wnd[0]").sendVKey 3
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
end function