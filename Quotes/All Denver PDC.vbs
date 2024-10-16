dim fileSystemObject
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
  dim row, jobNumber
  row = 0
  do while true
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").firstVisibleRow = row
    jobNumber = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(row,"JOBS")
    if jobNumber <> "0" then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell row,"LANGTEXT_BUTTON"
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressButtonCurrentCell

      session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").selectItem "0002","COLUMN1"
      if err.number = 0 then
        session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "0002","COLUMN1"
        session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "0002","COLUMN1"

        session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "Denver PDC" + vbCr + ""
      else
        err.clear
      end if
      session.findById("wnd[1]/tbar[0]/btn[8]").press
    end if

    if err.number <> 0 then
      if session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(row - 1,"ITOBJID") <> "" then
        err.clear
      else
        exit do
      end if
    end if
    
    row = row + 1
  loop
end sub