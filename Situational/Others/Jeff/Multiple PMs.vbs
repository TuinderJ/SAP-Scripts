dim PRODUCTION
PRODUCTION = true

function readIntervals()
  err.clear
  on error resume next
  redim serviceInterval(2, 0)
  i = 0
  x = 0

  do while true
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    dim intervalName
    intervalName = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
    if err.number <> 0 then
      err.clear
      exit do
    end if
    'check if there is a reefer interval
    if not truckHasReefer and inStr(1, intervalName, "RFPM") > 0 then
      truckHasReefer = true
      workOrder.add "Reefer", session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
      if not PRODUCTION then
        msgBox "Truck has reefer"
      end if
    end if
    'if PM
    if inStr(1, intervalName, "PM") > 0 and inStr(1, intervalName, "RFPM") = 0 then
      storeInterval()
    end if
    i = i + 1
  loop
  i = 0
  do while true
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    intervalName = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
    if err.number <> 0 then
      exit do
    end if
    'if OF
    if inStr(1, intervalName, "OF") > 0 then
      if checkIfDue(intervalName) = true then
        storeInterval()
      else
        workOrder.add "Oil", session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
      end if
    end if
    'if DOT, DRYR, DEFFI
    if _
    inStr(1, intervalName, "DOT") > 0 Or _
    inStr(1, intervalName, "DRYR") > 0 Or _
    inStr(1, intervalName, "DEFFI") > 0 _
    then
      if checkIfDue(intervalName) = true then
        if not PRODUCTION then
          msgBox("It decided the interval is due")
        end if
        storeInterval()
      end if
    end if
    i = i + 1
  loop
end function

function storeInterval()
  if not PRODUCTION then  
    msgBox "storing"
  end if
  redim Preserve serviceInterval(2, x)
  serviceInterval(0, x) = x + 1
  serviceInterval(1, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
  serviceInterval(2, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
  x = x + 1
end function

function checkIfDue(intervalName)
  result = false
  dim intervalDueDate, intervalMileage, unitofMeasure, truckCurrentMileage, reeferCurrentHours
  intervalDueDate = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")
  intervalMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
  unitofMeasure = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT_U")
  truckCurrentMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text
  reeferCurrentHours = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZREEFER_HRS").text
  if not PRODUCTION then
    msgBox("intervalDueDate: " & intervalDueDate & vbCr & "currentDate: " & date & vbCr & "intervalMileage: " & intervalMileage & vbCr & "truckCurrentMileage: " & truckCurrentMileage)
  end if
  if intervalDueDate <> "" then
    if instr(intervalName, "RFPM") > 0 then
      if not PRODUCTION then
        msgBox "Reefer interval"
      end if
      if reeferCurrentHours > intervalMileage then
        result = true
        if not PRODUCTION then
          msgBox "Due by reefer hours"
        end if
      end if
    else
      if isNumeric(right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) then
        if date + CInt(right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) > CDate(intervalDueDate) then
          result = true
          if not PRODUCTION then
            msgBox "Due by time 1"
          end if
        end if
      else
        if date + 90 > CDate(intervalDueDate) then
          result = true
          if not PRODUCTION then
            msgBox "Due by time 2"
          end if
        end if
      end if
    end if
  end if
  if result = false then
    on error resume next
    if unitofMeasure = "MI" then
      if cLng(intervalMileage) <> 0 then
        if CLng(intervalMileage) <= CLng(truckCurrentMileage) then
          result = true
          if not PRODUCTION then
            msgBox "Due by Miles"
          end if
        end if
      end if
    end if
  end if
  err.clear
  checkIfDue = result
end function

function makeJobs()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  for job = 0 to UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = serviceInterval(1, job) & " -" & serviceInterval(2, job)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = serviceInterval(1, job)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 5
  next
end function

function addLabor()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  for job = 0 to UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = serviceInterval(0, job)
    if inStr(1, serviceInterval(1, job),"PM") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM HD"
    end if
    if inStr(1, serviceInterval(1, job),"OF") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM OF"
    end if
    if inStr(1, serviceInterval(1, job),"DOT") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "10"
    end if
    if inStr(1, serviceInterval(1, job),"DRYR") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "1310056"
    end if
    if inStr(1, serviceInterval(1, job),"RFPM") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "9800050"
    end if
    if inStr(1, serviceInterval(1, job),"DEFFI") > 0 then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "4307002"
    end if
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
  next
end function

function readOrder()
  dim title, unit
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01").select
  title = session.findById("wnd[0]/titl").text
  title = replace(title, "&", "")
  workOrder.add "Customer", right(title, len(title) - 40)
  workOrder.add "RO", right(Left(title,37),8)
  unit = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txtIS_VLCACTDATA_ITEM-ZZUN").text
  customerUnit = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txtIS_VLCACTDATA_ITEM-VHCEX").text
  if unit <> "" then
    if customerUnit <> "" then
      workOrder.add "Unit", Left(unit, 3) & "-" & right(unit, len(unit) - 3) & " / " & customerUnit
    else
      workOrder.add "Unit", Left(unit, 3) & "-" & right(unit, len(unit) - 3)
    end if
  else
    workOrder.add "Unit", customerUnit
  end if
end function

function convertServiceIntervalstoJobs()
  for job = 0 to UBound(serviceInterval, 2)
    redim Preserve jobs(job)
    jobs(job) = serviceInterval(1, job) & " -" & serviceInterval(2, job)
  next
end function

function printoldSheetAndMakeNew()
  if isObject(objExcel) then
    if PRODUCTION = true then
      objWorkbook.Printout
    end if
    objWorkbook.Close false
    objExcel.workbooks.Close
    objExcel.quit
  end if

  Set objExcel = createObject("Excel.Application")
  Set objWorkbook = objExcel.workbooks.add()

  if PRODUCTION = false then
    objExcel.visible = true
    objExcel.windowState = -4137
  end if

  objExcel.displayAlerts = false
  objWorkbook.workSheets.item(1).pageSetup.centerHeader = workOrder.item("Customer")
end function

function addtoSheet()
  if toggle = "top" then
    rowtoStart = 1
    toggle = "Bottom"
  else
    rowtoStart = 24
  end if
  if rowtoStart = 1 then
    printoldSheetAndMakeNew()
  else
    if objWorkbook.workSheets.item(1).pageSetup.centerHeader <> workOrder.item("Customer") then
      printoldSheetAndMakeNew()
      rowtoStart = 1
    else
      toggle = "top"
    end if
  end if

  ' RO
  objExcel.columns("A:I").columnWidth = "9.1"
  objExcel.range("A" & rowtoStart, "C" & (rowtoStart + 1)).merge
  with objExcel.range("A" & rowtoStart)
    .value = workOrder.item("RO")
    .font.size = 18
    .horizontalAlignment = -4108
  end with

  ' Unit number
  objExcel.range("D" & rowtoStart, "F" & (rowtoStart + 1)).merge
  with objExcel.range("D" & rowtoStart)
    .value = workOrder.item("Unit")
    .font.size = 18
    .horizontalAlignment = -4108
  end with

  ' Last DOT
  objExcel.range("G" & rowtoStart, "G" & (rowtoStart + 1)).merge
  with objExcel.range("G" & rowtoStart)
    .value = "Last DOT:"
    .horizontalAlignment = -4152
  end with
  objExcel.range("H" & rowtoStart, "I" & (rowtoStart + 1)).merge
  with objExcel.range("H" & rowtoStart, "I" & (rowtoStart + 1)).borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Miles
  objExcel.range("A" & (rowtoStart + 2), "A" & (rowtoStart + 3)).merge
  with objExcel.range("A" & (rowtoStart + 2))
    .value = "Miles:"
    .font.size = 14
    .horizontalAlignment = -4152
  end with
  objExcel.range("B" & (rowtoStart + 2), "C" & (rowtoStart + 3)).merge
  with objExcel.range("B" & (rowtoStart + 2), "C" & (rowtoStart + 3)).borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Hours
  objExcel.range("A" & (rowtoStart + 4), "A" & (rowtoStart + 5)).merge
  with objExcel.range("A" & (rowtoStart + 4))
    .value = "Hours:"
    .font.size = 14
    .horizontalAlignment = -4152
  end with
  objExcel.range("B" & (rowtoStart + 4), "C" & (rowtoStart + 5)).merge
  with objExcel.range("B" & (rowtoStart + 4), "C" & (rowtoStart + 5)).borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Reefer Hours
  if truckHasReefer then
    objExcel.range("A" & (rowtoStart + 6), "B" & (rowtoStart + 6)).merge
    with objExcel.range("A" & (rowtoStart + 6))
      .value = "Reefer Hours:"
      .font.size = 14
      .horizontalAlignment = -4152
    end with
    objExcel.range("C" & (rowtoStart + 6), "D" & (rowtoStart + 6)).merge
    with objExcel.range("C" & (rowtoStart + 6), "D" & (rowtoStart + 6)).borders(9)
      .lineStyle = 1
      .weight = 2
      .colorIndex = -4105
    end with
  end if

  ' Fuel Filters
  objExcel.range("D" & (rowtoStart + 2), "E" & (rowtoStart + 3)).merge
  with objExcel.range("D" & (rowtoStart + 2))
    .value = "Fuel Filters?"
    .horizontalAlignment = -4152
    .VerticalAlignment = -4108
  end with
  objExcel.range("F" & (rowtoStart + 2), "F" & (rowtoStart + 3)).merge
  with objExcel.range("F" & (rowtoStart + 2), "F" & (rowtoStart + 3)).borders(7)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 2), "F" & (rowtoStart + 3)).borders(8)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 2), "F" & (rowtoStart + 3)).borders(9)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 2), "F" & (rowtoStart + 3)).borders(10)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with

  ' Oil Filters
  objExcel.range("D" & (rowtoStart + 4), "E" & (rowtoStart + 5)).merge
  with objExcel.range("D" & (rowtoStart + 4))
    .value = "Oil Filters?"
    .horizontalAlignment = -4152
    .VerticalAlignment = -4108
  end with
  objExcel.range("F" & (rowtoStart + 4), "F" & (rowtoStart + 5)).merge
  with objExcel.range("F" & (rowtoStart + 4), "F" & (rowtoStart + 5)).borders(7)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 4), "F" & (rowtoStart + 5)).borders(8)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 4), "F" & (rowtoStart + 5)).borders(9)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with
  with objExcel.range("F" & (rowtoStart + 4), "F" & (rowtoStart + 5)).borders(10)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  end with

  ' date Completed
  objExcel.range("G" & (rowtoStart + 2), "I" & (rowtoStart + 3)).merge
  with objExcel.range("G" & (rowtoStart + 2))
    .value = "date Completed:"
    .font.size = 18
    .horizontalAlignment = -4108
  end with
  objExcel.range("G" & (rowtoStart + 4), "I" & (rowtoStart + 5)).merge
  with objExcel.range("G" & (rowtoStart + 4), "I" & (rowtoStart + 5)).borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Notes Box
  with objExcel.range("A" & (rowtoStart + 7), "I" & (rowtoStart + 22)).borders(7)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("A" & (rowtoStart + 7), "I" & (rowtoStart + 22)).borders(8)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("A" & (rowtoStart + 7), "I" & (rowtoStart + 22)).borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("A" & (rowtoStart + 7), "I" & (rowtoStart + 22)).borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Jobs
  objExcel.range("A" & (rowtoStart + 7), "C" & (rowtoStart + 7)).merge
  with objExcel.range("A" & (rowtoStart + 7), "C" & (rowtoStart + 7)).borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("A" & (rowtoStart + 7), "C" & (rowtoStart + 7)).borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("A" & (rowtoStart + 7))
    .value = "Jobs"
    .horizontalAlignment = -4108
  end with
  iterator = 0
  for Each job in jobs
    with objExcel.range("A" & (rowtoStart + iterator + 8), "C" & (rowtoStart + iterator + 8))
      .merge
      with .borders(9)
        .lineStyle = 5
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
    end with
    objExcel.range("A" & (rowtoStart + iterator + 8)).value = job
    iterator = iterator + 1
  next
  with objExcel.range("A" & (rowtoStart + iterator + 8), "C" & (rowtoStart + iterator + 8)).borders(8)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' Repairs / Notes
  objExcel.range("D" & (rowtoStart + 7), "F" & (rowtoStart + 7)).merge
  with objExcel.range("D" & (rowtoStart + 7), "F" & (rowtoStart + 7)).borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("D" & (rowtoStart + 7))
    .value = "Repairs / Notes"
    .horizontalAlignment = -4108
  end with
  with objExcel.range("F" & (rowtoStart + 7), "F" & (rowtoStart + 22)).borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  end with

  ' If there's no oil change interval, add the mileage that it's due at
  if workOrder.item("Oil") <> "" and workOrder.item("Oil") <> vbCr then
    with objExcel.range("D" & (rowtoStart + 8), "F" & (rowtoStart + 8))
      .merge
      .value = "Oil change due at " & workOrder.item("Oil") & " miles."
    end with
  else
    with objExcel.range("F" & (rowtoStart + 4))
      .font.size = 24
      .horizontalAlignment = -4108
      .VerticalAlignment = -4108
      .value = "X"
    end with
  end if

  ' If there's a reefer, add the reefer hours that it's due at
  if truckHasReefer then
    with objExcel.range("D" & (rowtoStart + 9), "F" & (rowtoStart + 9))
      .merge
      .value = "Reefer service due at " & workOrder.item("Reefer") & " hours."
    end with
  end if

  ' Parts Need to Order
  objExcel.range("G" & (rowtoStart + 7), "I" & (rowtoStart + 7)).merge
  with objExcel.range("G" & (rowtoStart + 7), "I" & (rowtoStart + 7)).borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  end with
  with objExcel.range("G" & (rowtoStart + 7))
    .value = "Parts Need to Order"
    .horizontalAlignment = -4108
  end with

  ' Tires
  with objExcel.range("A" & (rowtoStart + 18))
    .value = "Tires"
    .horizontalAlignment = -4108
  end with
  for i = 2 to 3
    with objExcel.Cells(rowtoStart + 18, i)
      .Interior.colorIndex = 15
      with .borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
    end with
  next
  for i = 1 to 3
    with objExcel.Cells(rowtoStart + 19, i)
      .Interior.colorIndex = 15
      with .borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
    end with
  next
  for i = 1 to 3
    with objExcel.Cells(rowtoStart + 21, i)
      .Interior.colorIndex = 15
      with .borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
    end with
  next
  for i = 2 to 3
    with objExcel.Cells(rowtoStart + 22, i)
      .Interior.colorIndex = 15
      with .borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
      with .borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      end with
    end with
  next
end function

if not isObject(application) then
   Set SapGuiAuto  = Getobject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
end if
if not isObject(connection) then
   Set connection = application.Children(0)
end if
if not isObject(session) then
   Set session    = connection.Children(0)
end if
if isObject(wScript) then
   wScript.Connectobject session,     "on"
   wScript.Connectobject application, "on"
end if

on error resume next
session.findById("wnd[0]").maximize

dim truck()
dim serviceInterval()
dim jobs()
dim userInput, i, x, verifyTrucks, trucksNotFound, trucksFound, result, workOrder, objExcel, objWorkbook, toggle, truckHasReefer
Set workOrder = createObject("Scripting.Dictionary")
toggle = "top"
truckHasReefer = false

i = 0
' Ask the user which trucks they need
do while true
    userInput = InputBox("What is the unit number of the next truck?" & vbCr & "Don't use a -" & vbCr & "if you're done, leave it blank.", "add Trucks")
    if userInput = "" then
        exit do
    end if
    redim Preserve truck(i)
    truck(i) = userInput
    i = i + 1
loop

if truck(0) = "" then
  wScript.quit
end if

' Ask the user for verification
do while true
    i = 0
    verifyTrucks = ""
    for Each unit in truck
        verifyTrucks = verifyTrucks & vbCr & i + 1 & ": " & unit
        i = i + 1
    next

    if msgBox("Are all of these entered properly?" & verifyTrucks, vbYesNo, "Verify") =  vbNo then
        i = InputBox("What number do you need to change?" & "if you need to cancel the whole thing, leave blank." & verifyTrucks, "Change Which One")
        if i = "" then
            wScript.quit
        end if
        i = i - 1
        userInput = InputBox("What is the new number for " & truck(i) & "?" & vbCr & "Leave blank to remove it.", "New Unit number")
        if userInput <> "" then
            truck(i) = userInput
        else
            if i = UBound(truck) then
                redim Preserve truck(i - 1)
            end if
            if i <= UBound(truck)then
                do Until i => UBound(truck)
                    truck(i) = truck(i + 1)
                    i = i + 1
                loop
                redim Preserve truck(i - 1)
            end if
        end if
    else
        exit do
    end if
loop

for each unit in truck
  on error resume next
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unit
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
  
  do
    if session.findById("wnd[0]/sbar").text = "No vehicles could be selected" then
      trucksNotFound = trucksNotFound & vbCr & unit
      trucksFound = trucksFound & vbCr
      exit do
    else
      trucksFound = trucksFound & vbCr & unit
      session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
      session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "74247"
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "7039"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      err.clear

      'Header
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-ZZORD_STATUS").key = "ON-SITE/MOBILE REPAIR"
      session.findById("wnd[0]").sendVKey 0

      readIntervals()
      makeJobs()
      addLabor()
      session.findById("wnd[0]").sendVKey 11
      session.findById("wnd[0]/tbar[1]/btn[13]").press
      convertServiceIntervalstoJobs()
      workOrder.add "Jobs", jobs
      readOrder()
      session.findById("wnd[0]/tbar[0]/btn[3]").press
    end if
  loop while false
  on error goto 0
  addtoSheet()
  workOrder.removeAll
  redim jobs(0)
next

if isObject(objExcel) then
  if PRODUCTION = true then
    objWorkbook.Printout
  end if
  objWorkbook.Close false
  objExcel.workbooks.Close
  objExcel.quit
end if

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press

if trucksNotFound <> "" then
    msgBox("These trucks were not found in SAP." & trucksNotFound)
end if
if trucksFound <> "" then
    msgBox("These trucks were created successfully." & trucksFound)
end if

Set objExcel = nothing
Set objWorkbook = nothing