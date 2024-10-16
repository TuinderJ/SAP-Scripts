dim fileSystemObject
set fileSystemObject = createObject("Scripting.FileSystemObject")
include("Z:\utilities.vbs")

dim userInput
set userInput = createObject("Scripting.Dictionary")

main()
set fileSystemObject = nothing

sub include (file)
	'Create objects for opening text file
	set fso = createObject("Scripting.FileSystemObject")
	set textFile = fso.openTextFile(file, 1)

	'Execute content of file.
	executeGlobal textFile.readAll

	'CLose file
	textFile.close

	'Clean up
	set fso = nothing
	set textFile = nothing
end sub

sub main()
  getUserInput()
  createRO()
end sub

sub getUserInput()
  userInput.add "invoiceNumber", inputBox("What is the invoice number for this ZP50?", "Invoice Number")
  if userInput.item("invoiceNumber") = "" then
    wScript.quit
  end if

  dim parts()
  dim i
  i = 0
  do while true
    redim preserve parts(1, i)
    parts(0, i) = inputBox("What is the part number?", "Part Number")
    if parts(0, i) = "" then
      exit do
    end if

    parts(1, i) = inputBox("What is the quantity?", "Quantity")
    if parts(1, i) = "" then
      wScript.quit
    end if
    i = i + 1
  loop

  userInput.add "parts", parts

  userInput.add "customerNumber", inputBox("What is the account number this is being billed to?", "Customer", "90038")
  if userInput.item("customerNumber") = "" then
    wScript.quit
  end if

  userInput.add "billedAtCost", (msgBox("Is this being billed at cost?", vbYesNo, "Cost") = vbYes)
end sub

sub createRO()
  goToTCode("/DBM/ORDER01")

  ' Fill out order creation form
  session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP50"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = userInput.item("customerNumber")
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
  session.findById("wnd[0]/tbar[1]/btn[13]").press

  ' Fill out header
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2065/subSUBSCREEN_2065:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3201/txt/DBM/VBAK_COM-BSTNK").text = "ZP50 FOR " & userInput.item("invoiceNumber")
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2065/subSUBSCREEN_2065:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3201/cmb/DBM/VBAK_COM-KDGRP").key = "04"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2070/subSUBSCREEN_2070:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:SAPLZZMM031_PARTS:2010/txt/DBM/VBAK_COM-ZZDBM_ORG_INV").text = userInput.item("invoiceNumber")
  session.findById("wnd[0]").sendVKey 0

  ' Parts tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "L"

  for i = 0 to uBound(userInput.item("parts"), 2) - 1
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = userInput("parts")(0, i)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = userInput("parts")(1, i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnCORE1").press
  next

  for i = 0 to uBound(userInput.item("parts"), 2) - 1
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell i,"KBETM",session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(i,"ZZWAVWR")
  next
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressEnter
  session.findById("wnd[0]/tbar[0]/btn[11]").press
  session.findById("wnd[0]/tbar[1]/btn[45]").press
  session.findById("wnd[1]/usr/cntlSELECTION_SCREEN1/shellcont/shell").selectAll
  session.findById("wnd[1]/tbar[0]/btn[2]").press
  session.findById("wnd[0]/tbar[1]/btn[13]").press
  session.findById("wnd[0]/tbar[1]/btn[39]").press
  session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectAll
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/tbar[1]/btn[13]").press
  session.findById("wnd[1]/usr/btnBUTTON_1").press
end sub