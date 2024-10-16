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
  dim row, row2, itemNumber, itemNumber2, reqDate, po, shouldContinue
  goToTCode("ME5A")
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn "BANFN"
  session.findById("wnd[0]/tbar[1]/btn[28]").press

  ' for row = 1 to 5
  row = 1
  do while true
    ' reqDate = cDate(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(row,"BADAT"))
    ' if reqDate >= cDate("01/01/2024") then
      ' exit do
    ' end if
    shouldContinue = false
    ' po = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(row,"EBELN")
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell row,"BANFN"
    if msgBox("Delete this one?", vbYesNo, "Delete") = vbYes then
      shouldContinue = true
    end if

    if shouldContinue then
      itemNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(row,"BNFPO")
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell

      session.findById("wnd[0]/tbar[1]/btn[7]").press
      'Find the item number
      row2 = 0
      do while true
        session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").firstVisibleRow = row2
        itemNumber2 = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").getCellValue(row2,"BNFPO")
        ' msgBox("Item Number Searching For:" & vbCr & itemNumber & vbCr & vbCr & "Item Number Found:" & vbCr & itemNumber2 & vbCr & vbCr & "Row searching:" & vbCr & row2)
        if itemNumber = itemNumber2 then
          exit do
        end if
        row2 = row2 + 1
      loop


      'After you found the item number, delete it
      session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").selectedRows = cStr(row2)
      session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressToolbarButton "&MEREQDELETE"
      session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
      session.findById("wnd[0]/tbar[0]/btn[11]").press
      session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
    end if
  ' next
    row = row + 1
  loop

  if msgBox("Would you like to repeat?", vbYesNo, "Repeat") = vbYes then
    main()
  end if
end sub