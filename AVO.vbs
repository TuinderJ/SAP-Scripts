dim fileSystemObject
set fileSystemObject = createObject("Scripting.FileSystemObject")
include(fileSystemObject.getAbsolutePathName(".") & "\utilities.vbs")
set fileSystemObject = nothing


main()

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

dim vin, avo

sub main()
  askForUserInput()
  createOrder()
end sub

sub askForUserInput()
  if vin = "" then
    vin = inputBox("What is the vin?", "VIN")
  end if
  if vin = "" then
    wScript.quit
  end if

  if avo = "" then
    avo = inputBox("What is the AVO number?", "AVO")
  end if
  if avo = "" then
    wScript.quit
  end if
end sub

sub createOrder()
  goToTCode("/DBM/ORDER01")
  ' Fill out the data
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "100000"
  session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP21"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VHVIN").text = "*" & vin
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
  session.findById("wnd[0]/tbar[1]/btn[13]").press
  msgBox "Select the right one.",, "Validate"
  
  ' Partner selection
  session.findById("wnd[1]/tbar[0]/btn[71]").press
  session.findById("wnd[2]/usr/txtRSYSF-STRING").text = "006058550"
  session.findById("wnd[2]").sendVKey 0
  session.findById("wnd[3]/usr/lbl[4,2]").setFocus
  session.findById("wnd[3]").sendVKey 2
  session.findById("wnd[1]").sendVKey 2
  session.findById("wnd[1]/tbar[0]/btn[0]").press

  ' Header
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").setFocus
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "931"
  session.findById("wnd[0]").sendVKey 0

  ' Parts Tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "L"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text = "AVO " & avo
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "597915:PB"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = "2"
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]").sendVKey 0
  
  ' Save, release, pick
  session.findById("wnd[0]/tbar[0]/btn[11]").press
  session.findById("wnd[0]/tbar[1]/btn[37]").press
  session.findById("wnd[0]/tbar[1]/btn[45]").press
  session.findById("wnd[1]/tbar[0]/btn[2]").press
  session.findById("wnd[0]/tbar[1]/btn[13]").press
  session.findById("wnd[0]").sendVKey 0
end sub
