dim fileSystemObject
set fileSystemObject = createObject("Scripting.FileSystemObject")
' include(fileSystemObject.getAbsolutePathName(".") & "\utilities.vbs")
include("Z:\utilities.vbs")

dim deletedParts()
redim deletedParts(0)
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
  openZZREQ()
  goThroughVendors()
  makeCSVFile()
  msgBox "Finished",, "Complete"
end sub

sub openZZREQ()
  goToTCode("ZZREQ")
  ' Fill out fields
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1401"
  if weekday(date) = 2 then
    previousDate = date - 3
  else
    previousDate = date - 1
  end if
  session.findById("wnd[0]/usr/btn%_S_BSART_%_APP_%-VALU_PUSH").press
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ZMDI"
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ZSOR"
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "ZRGT"
  session.findById("wnd[1]/tbar[0]/btn[8]").press
  session.findById("wnd[0]/usr/ctxtS_BADAT-LOW").text = previousDate
  session.findById("wnd[0]/usr/ctxtS_BADAT-HIGH").text = date
  session.findById("wnd[0]/tbar[1]/btn[8]").press
end sub

sub goThroughVendors()
  for i = 0 to 1
    dim row, message

    row = 0
    do
      if row >= session.findById("wnd[0]/shellcont[" & i & "]/shell").rowCount - 1 then
        exit do
      end if
      dim minimumValue, currentValue, minimumQty, currentQty
      minimumValue = session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"ZZ_DOLLAR")
      currentValue = session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"NETPR")
      minimumQty = session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"ZZ_QTY")
      currentQty = session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"MENGE")

      if uCase(minimumValue) = "NO MINS" then
        minimumValue = "0"
      end if
      if uCase(minimumQty) = "NO MINS" then
        minimumQty = "0"
      end if

      message = _
        "Vendor: " & session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"EMNFR") & _
        vbCr & _
        session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"NAME1")
      if cDbl(minimumValue) <> 0 then
        message = message & _
          vbCr & vbCr
        if cDbl(minimumValue) > cDbl(currentValue) then
          message = message & _
            "Order Value - Below Minimum" & _
            vbCr & _
            "---------------------------------------"
        else
          message = message & _
            "Order Value" & _
            vbCr & _
            "----------------"
        end if
        message = message & _
          vbCr & _
          "Minimum: $" & formatNumber(minimumValue, 2) & _
          vbCr & _
          "Current: $" & formatNumber(currentValue, 2)
      end if
      if cDbl(minimumQty) > 0 then
        message = message & _
          vbCr & vbCr
        if cDbl(minimumQty) > cDbl(currentQty) then
          message = message & _
            "Order Qty - Below Minimum" & _
            vbCr & _
            "------------------------------------"
        else
          message = message & _
            "Order Qty" & _
            vbCr & _
            "-------------"
        end if
        message = message & _
          vbCr & _
          "Minimum: " & minimumQty & _
          vbCr & _
          "Current: " & currentQty
      end if
      dim exclude
      exclude = false
      
      if session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"EMNFR") = "IMS" then
        exclude = true
      end if
      if session.findById("wnd[0]/shellcont[" & i & "]/shell").getCellValue(row,"EMNFR") = "FRD" then
        exclude = true
      end if
      if not exclude then
        session.findById("wnd[0]/shellcont[" & i & "]/shell").setCurrentCell row,"EMNFR"
        session.findById("wnd[0]/shellcont[" & i & "]/shell").clickCurrentCell
        dim shouldDisplayMessage
        if cDbl(currentValue) < cDbl(minimumValue) or cDbl(currentQty) < cDbl(minimumQty) then
          shouldDisplayMessage = true
        else
          shouldDisplayMessage = false
        end if
        if goThroughReqs(message, shouldDisplayMessage) = vbNo then
          row = row + 1
        end if
      else
        row = row + 1
      end if
    loop while true
  next
end sub

function goThroughReqs(message, shouldDisplayMessage)
  goThroughReqs = vbNo
  dim rowCount, deleteCount, highlightedRows, requestsAreOnlyAutoGenerated
  highlightedRows = ""
  deleteCount = 0
  rowCount = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount
  requestsAreOnlyAutoGenerated = true
  on error resume next
  for row = 0 to rowCount - 1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = row - 4
    err.clear
    dim createdBy
    createdBy = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"ERNAM")
    if markLineForDelete(row) then
      deleteCount = deleteCount + 1
      if markLineForDelete then
        dim part
        set part = createObject("Scripting.Dictionary")
        part.add "partNumber", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"MATNR")
        part.add "description", replace(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"TXZ01"), ",", "-")
        part.add "qtyRequested", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"MENGE")
        part.add "openPoQty", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"OPEN_PO_QTY")
        part.add "roundingValue", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"BSTRF")
        part.add "reorderPoint", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"MINBE")
        part.add "qtySold", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"TTL_HIGH")
        part.add "lastPurchaseDate", session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"AEDAT")
        redim preserve deletedParts(uBound(deletedParts) + 1)
        set deletedParts(uBound(deletedParts)) = part
      end if
    end if

    if createdBy <> "Order Group" and createdBy <> "BATCH_USER" then
      requestsAreOnlyAutoGenerated = false
    end if
  next
  if deleteCount > 0 then
    pressDeleteButton()
    if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").rowCount = 0 then
      goThroughReqs = vbYes
    end if
    if requestsAreOnlyAutoGenerated then
      goThroughReqs = vbYes
    end if
  elseif requestsAreOnlyAutoGenerated and shouldDisplayMessage then
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = highlightedRows
    goThroughReqs = msgbox("WAIT!!!!!!!!!!!" & vbCr & message & vbCr & vbCr & "Do you want to delete the entire order?", vbYesNo, "Manual Review")
    if goThroughReqs = vbYes then
      session.findById("wnd[0]/tbar[1]/btn[19]").press
      pressDeleteButton()
    end if
  end if
  goBackNScreens 1
end function

function markLineForDelete(row)
  dim createdBy, shouldContinue
  shouldContinue = false
  createdBy = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"ERNAM")
  markLineForDelete = false

  if createdBy = "Order Group" or createdBy = "BATCH_USER" then
    shouldContinue = true
  end if
  if not shouldContinue then
    exit function
  end if

  dim pastTwelveMonths(11)
  for i = 0 to 11
    pastTwelveMonths(i) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"MONTH" & i + 1)
  next
  if session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(row,"TTL_HIGH") = "" then
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").modifyCheckbox row,"CHECK_PR",true
    markLineForDelete = true
    exit function
  else
    if _
      pastTwelveMonths(0) = "0" and _
      pastTwelveMonths(1) = "0" and _
      pastTwelveMonths(2) = "0" and _
      pastTwelveMonths(3) = "0" _
      then
      session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").modifyCheckbox row,"CHECK_PR",true
      markLineForDelete = true
      exit function
    end if
    for i = 0 to 11
      dim zeroCount
      zeroCount = 0
      if pastTwelveMonths(i) = "0" then
        zeroCount = zeroCount + 1
      end if
      if zeroCount > 10 then
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").modifyCheckbox row,"CHECK_PR",true
        markLineForDelete = true
      end if
    next
  end if
end function

sub makeCSVFile()
  if uBound(deletedParts) = 0 then
    exit sub
  end if

  dim csvFilePath, csvColumns
  csvFilePath = WScript.CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Deleted Items " & replace(date, "/", "-") & ".csv"
  csvColumns = "Part Number, Description, Qty Requested, Open PO Qty, Rounding Value, Reorder Point, Qty Sold, Last Purchase Date"
  Set csvFile = fileSystemObject.createTextFile(csvFilePath, true)
  csvFile.Write csvColumns
  csvFile.Writeline

  for row = 1 to uBound(deletedParts)
    dim rowData
    rowData = deletedParts(row).item("partNumber") & "," & _
      deletedParts(row).item("description") & "," & _
      deletedParts(row).item("qtyRequested") & "," & _
      deletedParts(row).item("openPoQty") & "," & _
      deletedParts(row).item("roundingValue") & "," & _
      deletedParts(row).item("reorderPoint") & "," & _
      deletedParts(row).item("qtySold") & "," & _
      deletedParts(row).item("lastPurchaseDate")
    csvFile.Write rowData
    csvFile.Writeline
  next

end sub

sub pressDeleteButton()
  session.findById("wnd[0]/tbar[1]/btn[18]").press
  ' msgBox "It will press delete here."
  on error resume next
  session.findById("wnd[1]/tbar[0]/btn[8]").press
  err.clear
end sub