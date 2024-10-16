dim fileSystemObject
set fileSystemObject = createObject("Scripting.FileSystemObject")
include("Z:\utilities.vbs")

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
  on error resume next
  dim orderType, currentMode
  orderType = split(split(session.findById("wnd[0]/titl").text,":")(0)," ")(0)
  currentMode = split(split(session.findById("wnd[0]/titl").text,":")(0)," ")(2)
  
  session.findById("wnd[1]/tbar[0]/btn[12]").press
  session.findById("wnd[2]/usr/btnBUTTON_1").press
  err.clear
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  if currentMode = "Display" then
    session.findById("wnd[0]/tbar[1]/btn[13]").press
  end if

  dim file, line
  set file = fileSystemObject.openTextFile("C:\Users\tuinderj\OneDrive - Rush Enterprises\Documents\Quick Add.txt")
  do while not file.atEndOfStream
    line = split(file.readLine())
    if inStr(line(1), ":") then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = line(1)
    else
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = line(1) & "*"
    end if
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = line(0)
    session.findById("wnd[0]").sendVKey 0
    dim qtyOnHand
    qtyOnHand = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-VRFMG_LGORT").text
    if cInt(qtyOnHand) < cInt(line(0)) then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-JOBS").text = "1"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[1]/usr/txt/DBM/JOB_COM-DESCR1").text = "NOT IN STOCK"
      session.findById("wnd[1]/tbar[0]/btn[0]").press
    end if
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnCORE1").press
    session.findById("wnd[1]").sendVKey 0
    session.findById("wnd[1]").sendVKey 0
  loop
end sub