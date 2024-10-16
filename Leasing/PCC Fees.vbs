dim poLines()
redim poLines(2, 0)
dim trucks()

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

askForUserInput()
createPO()
getPONumberAndMIGO()

function askForUserInput()
  addInvoices()

  dim verifyMessage
  verifyMessage = "Is all of this information correct?" & vbCr & vbCr

  for i = 0 to uBound(poLines, 2)
    verifyMessage = verifyMessage + poLines(0, i) & " - " & poLines(1, i) & " - " & poLines(2, i) & vbCr 
  next
  
  if msgBox(verifyMessage, vbYesNo, "Verify") = vbNo then
    WScript.Quit
  end if

end function

function addInvoices()
  dim invoiceNumber, vin, price, poLineCount
  poLineCount = uBound(poLines, 2)
  invoiceNumber = inputBox("What is the next invoice number?", "Invoice Number")
  if invoiceNumber = "" then
    exit function
  end if
  vin = inputBox("What is the VIN provided?", "VIN")
  price = inputBox("What is the price of the invoice?", "Price", "45")
  
  if poLines(0, 0) = "" then
    poLines(0, 0) = invoiceNumber
    poLines(1, 0) = vin
    poLines(2, 0) = price
  else
    redim preserve poLines(2, poLineCount + 1)
    poLines(0, poLineCount + 1) = invoiceNumber
    poLines(1, poLineCount + 1) = vin
    poLines(2, poLineCount + 1) = price
  end if
  
  addInvoices()

end function

function createPO()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME21N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  
  dim itteration, vendorNumber
  itteration = findItteration()
  vendorNumber = "205174"
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = vendorNumber
  session.findById("wnd[0]").sendVKey 0
  
  for i = 0 to uBound(poLines, 2)
    itteration = findItteration()
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2," & i & "]").text = "F"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[4," & i & "]").text = poLines(0, i)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[5," & i & "]").text = "1"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[6," & i & "]").text = "EA"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7," & i & "]").text = poLines(2, i)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[9," & i & "]").text = "1111"
  next
  session.findById("wnd[0]").sendVKey 0

  ' Decide what item to click on
  dim a, b, c, d
  a = 0
  b = " "
  c = " "
  d = " "
  for i = 0 to uBound(poLines, 2)
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
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = "531030"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = branch & "00"

    ' Select order number (which truck)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-AUFNR").setFocus
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB016/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]").text = "*" & poLines(1, i)
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
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
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  dim purchaseOrderNumber
  purchaseOrderNumber = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & findItteration() & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/txtMEPO_TOPLINE-EBELN").text
  session.findById("wnd[0]/tbar[0]/btn[11]").press

  session.findById("wnd[0]/tbar[0]/okcd").text = "/NMIGO"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  ' Start loop
  for i = 0 to uBound(poLines, 2)
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = purchaseOrderNumber
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_FIRSTLINE:SAPLMIGO:0011/ctxtGODEFAULT_TV-BWART").text = "101"
    session.findById("wnd[0]").sendVKey 0

    invoiceNumber = session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-MAKTX[2,0]").text
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").selected = true
    session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = invoiceNumber
    session.findById("wnd[0]/tbar[1]/btn[23]").press
  next

  session.findById("wnd[0]").sendVKey 3
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
end function