dim row, unitNumber, skip
dim parts()
row = 0
skip = false

function goToStos()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZONOREP"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7039"
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]").sendVKey 8
  session.findById("wnd[0]/tbar[1]/btn[33]").press
  session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").setCurrentCell 5,"TEXT"
  session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "5"
  session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
end function


function goToFirstTicketAndGrabPo()
  on error resume next
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell row,"BEDNR"
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session.findById("wnd[1]").sendVKey 0
  err.clear
  unitNumber = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text
  if msgBox("Do you want to move forward with this?", vbYesNo, "Keep Going") = vbNo then
    skip = true
    session.findById("wnd[0]").sendVKey 3
  end if
end function

function goToNextTicketAndGrabPo()
  on error resume next
  dim oldTicket
  oldTicket = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"BEDNR")
  do while true
    row = row + 1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell row,"BEDNR"
    if err.number <> 0 then
      goToNextTicketAndGrabPo = false
      exit function
    end if
    if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"BEDNR") <> oldTicket then
      exit do
    end if
  loop

  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session.findById("wnd[1]").sendVKey 0
  unitNumber = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text
  if msgBox("Do you want to move forward with this?", vbYesNo, "Keep Going") = vbNo then
    skip = true
    session.findById("wnd[0]").sendVKey 3
  end if
end function

function grapParts()
  if skip = true then
    exit function
  end if
  on error resume next
  dim currentPartRow, i
  currentPartRow = 0
  i = 0
  do until err.number <> 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell currentPartRow + 1,"ITOBJID"
    if session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(currentPartRow,"ITOBJID") <> "" then
      redim preserve parts(1, i)
      parts(0, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(i,"ITOBJID")
      parts(1, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(currentPartRow,"ZMENG")
      i = i + 1
    end if
    currentPartRow = currentPartRow + 1
  loop
end function

function goToVehicleOnSecondScreen()
  if skip = true then
    exit function
  end if
  session2.findById("wnd[0]").maximize
  session2.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session2.findById("wnd[0]/tbar[0]/btn[0]").press
  session2.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = replace(unitNumber, "-", "")
  session2.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
  session2.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
  session2.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:/DBM/SAPLVM19:2000/cntlG_CONTAINER/shellcont/shell").currentCellColumn = "VBELN"
  session2.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:/DBM/SAPLVM19:2000/cntlG_CONTAINER/shellcont/shell").selectedRows = "0"

  ' First RO
  if msgBox("Would you like to make a new RO?", vbYesNo, "New RO") = vbYes then
    makeNewRo()
  else
    msgBox "Click on the RO that you want to put the parts on.", vbOk, "WAIT!!!!!"
    session2.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:/DBM/SAPLVM19:2000/btnBUTTON1").press
    session2.findById("wnd[0]/tbar[1]/btn[13]").press
  end if
end function

function makeNewRo()
  if skip = true then
    exit function
  end if
  dim jobDescription
  session2.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "7039"
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").setFocus
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").caretPosition = 4
  session2.findById("wnd[0]/tbar[1]/btn[13]").press
  session2.findById("wnd[1]/tbar[0]/btn[0]").press
  session2.findById("wnd[0]").sendVKey 0
  ' Miles
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
  ' Hours
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
  session2.findById("wnd[0]").sendVKey 0
  ' Job
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  jobDescription = inputBox("What would you like the job to be called?", "Job Description", "PARTS")
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = jobDescription
  session2.findById("wnd[0]").sendVKey 0
  session2.findById("wnd[0]/tbar[0]/btn[11]").press
end function

function migoAndBillParts()
  if skip = true then
    exit function
  end if
  ' MIGO
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell 0,"ZZMIGO"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressButtonCurrentCell
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").selected = true
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/chkGOITEM-TAKE_IT[3,0]").setFocus
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/subSUB_BUTTONS:SAPLMIGO:0210/btnOK_TAKE_VALUE").press
  session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-LFSNR").text = unitNumber
  session.findById("wnd[0]/tbar[1]/btn[23]").press
  session.findById("wnd[0]").sendVKey 3

  ' Parts
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  ' Enter parts in
  on error resume next
  for i = 0 to uBound(parts,2)
    session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = parts(0,i)
    session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = parts(1,i)
    session2.findById("wnd[0]").sendVKey 0
    session2.findById("wnd[0]").sendVKey 0
    session2.findById("wnd[1]").sendVKey 0
    session2.findById("wnd[1]/usr/btnCORE1").press
  next
  session2.findById("wnd[0]/tbar[0]/btn[11]").press
end function









If Not IsObject(application) Then
  Set SapGuiAuto  = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
  Set connection = application.Children(0)
End If
If Not IsObject(session) Then
  Set session    = connection.Children(0)
  Set session2    = connection.Children(1)
End If
If IsObject(WScript) Then
  WScript.ConnectObject session,     "on"
  WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize


goToStos()
goToFirstTicketAndGrabPo()
goToVehicleOnSecondScreen()
grapParts()
migoAndBillParts()
skip = false

do while msgBox("Do you want to keep going?", vbYesNo, "Keep Going") = vbYes
  goToNextTicketAndGrabPo()
  goToVehicleOnSecondScreen()
  grapParts()
  migoAndBillParts()
  skip = false
loop