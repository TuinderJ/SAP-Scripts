Dim PRODUCTION
PRODUCTION = True

Function readIntervals()
  Err.Clear
  On Error Resume Next
  Redim serviceInterval(2, 0)
  i = 0
  x = 0

  Do While True
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    If Err.Number <> 0 Then
      Err.Clear
      Exit Do
    End If
    'If PM
    If inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"PM") = 1 Then
      storeInterval()
    End If
    i = i + 1
  Loop
  i = 0
  Do While True
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    If Err.Number <> 0 Then
        Exit Do
    End If
    'If OF
    If inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"OF") = 1 Then
      If checkIfDue() = True Then
        storeInterval()
      Else
        workOrder.Add "Oil", session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
      End If
    End If
    'If DOT, DRYR, RFPM, DEFFI
    If _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DOT") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DRYR") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"RFPM") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DEFFI") = 1 _
    Then
      If checkIfDue() = True Then
        ' MsgBox("It decided the interval is due")
        storeInterval()
      End If
    End If
    i = i + 1
  Loop
End Function

Function storeInterval()
  Redim Preserve serviceInterval(2, x)
  serviceInterval(0, x) = x + 1
  serviceInterval(1, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
  serviceInterval(2, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
  x = x + 1
End Function

Function checkIfDue()
  result = False
  Dim intervalDueDate, intervalMileage, truckCurrentMileage
  intervalDueDate = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")
  intervalMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
  truckCurrentMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text
  ' MsgBox("intervalDueDate: " & intervalDueDate & vbCr & "currentDate: " & Date & vbCr & "intervalMileage: " & intervalMileage & vbCr & "truckCurrentMileage: " & truckCurrentMileage)
  ' MsgBox("Date check: " & Date + CInt(Right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) & vbCr & "Interval: " & CDate(intervalDueDate))
  If intervalDueDate <> "" Then
    If Date + CInt(Right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) > CDate(intervalDueDate) Then
      result = True
    End If
  End If
  If result = False Then
    If _
    "" = intervalMileage Or _
    0 = CLng(intervalMileage) _
    Then
    Else
      If CLng(intervalMileage) <= CLng(truckCurrentMileage) Then
        result = True
      End If
    End If
  End If
  checkIfDue = result
End Function

Function makeJobs()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  For job = 0 to UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = serviceInterval(1, job) & " -" & serviceInterval(2, job)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = serviceInterval(1, job)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 5
  Next
End Function

Function addLabor()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  For job = 0 To UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = serviceInterval(0, job)
    If inStr(1, serviceInterval(1, job),"PM") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM HD"
    End If
    If inStr(1, serviceInterval(1, job),"OF") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM OF"
    End If
    If inStr(1, serviceInterval(1, job),"DOT") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "10"
    End If
    If inStr(1, serviceInterval(1, job),"DRYR") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "1310056"
    End If
    If inStr(1, serviceInterval(1, job),"RFPM") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "9800050"
    End If
    If inStr(1, serviceInterval(1, job),"DEFFI") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "4307002"
    End If
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
  Next
End Function

Sub readOrder()
  Dim title, unit
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01").select
  title = session.findById("wnd[0]/titl").text
  title = replace(title, "&", "")
  workOrder.add "Customer", Right(title, len(title) - 40)
  workOrder.add "RO", Right(Left(title,37),8)
  unit = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txtIS_VLCACTDATA_ITEM-ZZUN").text
  customerUnit = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txtIS_VLCACTDATA_ITEM-VHCEX").text
  if unit <> "" Then
    if customerUnit <> "" Then
      workOrder.add "Unit", Left(unit, 3) & "-" & Right(unit, len(unit) - 3) & " / " & customerUnit
    Else
      workOrder.add "Unit", Left(unit, 3) & "-" & Right(unit, len(unit) - 3)
    End If
  Else
    workOrder.add "Unit", customerUnit
  End If
End Sub

Sub convertServiceIntervalsToJobs()
  For job = 0 to UBound(serviceInterval, 2)
    Redim Preserve jobs(job)
    jobs(job) = serviceInterval(1, job) & " -" & serviceInterval(2, job)
  Next
End Sub

Sub printOldSheetAndMakeNew()
  If IsObject(objExcel) Then
    If PRODUCTION = true Then
      objWorkbook.PrintOut
    End If
    objWorkbook.Close False
    objExcel.workbooks.Close
    objExcel.Quit
  End If

  Set objExcel = CreateObject("Excel.Application")
  Set objWorkbook = objExcel.Workbooks.Add()

  If PRODUCTION = False Then
    objExcel.visible = True
    objExcel.WindowState = -4137
  End If

  objExcel.DisplayAlerts = False
  objWorkbook.WorkSheets.Item(1).PageSetup.CenterHeader = workOrder.item("Customer")
End Sub

Sub addToSheet()
  If toggle = "Top" Then
    rowToStart = 1
    toggle = "Bottom"
  Else
    rowToStart = 24
  End If
  If rowToStart = 1 Then
    printOldSheetAndMakeNew()
  Else
    If objWorkbook.WorkSheets.Item(1).PageSetup.CenterHeader <> workOrder.item("Customer") Then
      printOldSheetAndMakeNew()
      rowToStart = 1
    Else
      toggle = "Top"
    End If
  End If

  ' RO
  objExcel.Columns("A:I").ColumnWidth = "9.1"
  objExcel.Range("A" & rowToStart, "C" & (rowToStart + 1)).Merge
  With objExcel.Range("A" & rowToStart)
    .Value = workOrder.item("RO")
    .Font.Size = 18
    .HorizontalAlignment = -4108
  End With

  ' Unit Number
  objExcel.Range("D" & rowToStart, "F" & (rowToStart + 1)).Merge
  With objExcel.Range("D" & rowToStart)
    .Value = workOrder.item("Unit")
    .Font.Size = 18
    .HorizontalAlignment = -4108
  End With

  ' Last DOT
  objExcel.Range("G" & rowToStart, "G" & (rowToStart + 1)).Merge
  With objExcel.Range("G" & rowToStart)
    .Value = "Last DOT:"
    .HorizontalAlignment = -4152
  End With
  objExcel.Range("H" & rowToStart, "I" & (rowToStart + 1)).Merge
  with objExcel.Range("H" & rowToStart, "I" & (rowToStart + 1)).Borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Miles
  objExcel.Range("A" & (rowToStart + 2), "A" & (rowToStart + 3)).Merge
  With objExcel.Range("A" & (rowToStart + 2))
    .Value = "Miles:"
    .Font.Size = 14
    .HorizontalAlignment = -4152
  End With
  objExcel.Range("B" & (rowToStart + 2), "C" & (rowToStart + 3)).Merge
  With objExcel.Range("B" & (rowToStart + 2), "C" & (rowToStart + 3)).Borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Hours
  objExcel.Range("A" & (rowToStart + 4), "A" & (rowToStart + 5)).Merge
  With objExcel.Range("A" & (rowToStart + 4))
    .Value = "Hours:"
    .Font.Size = 14
    .HorizontalAlignment = -4152
  End With
  objExcel.Range("B" & (rowToStart + 4), "C" & (rowToStart + 5)).Merge
  with objExcel.Range("B" & (rowToStart + 4), "C" & (rowToStart + 5)).Borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Fuel Filters
  objExcel.Range("D" & (rowToStart + 2), "E" & (rowToStart + 3)).Merge
  With objExcel.Range("D" & (rowToStart + 2))
    .Value = "Fuel Filters?"
    .HorizontalAlignment = -4152
    .VerticalAlignment = -4108
  End With
  objExcel.Range("F" & (rowToStart + 2), "F" & (rowToStart + 3)).Merge
  With objExcel.Range("F" & (rowToStart + 2), "F" & (rowToStart + 3)).Borders(7)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 2), "F" & (rowToStart + 3)).Borders(8)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 2), "F" & (rowToStart + 3)).Borders(9)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 2), "F" & (rowToStart + 3)).Borders(10)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With

  ' Oil Filters
  objExcel.Range("D" & (rowToStart + 4), "E" & (rowToStart + 5)).Merge
  With objExcel.Range("D" & (rowToStart + 4))
    .Value = "Oil Filters?"
    .HorizontalAlignment = -4152
    .VerticalAlignment = -4108
  End With
  objExcel.Range("F" & (rowToStart + 4), "F" & (rowToStart + 5)).Merge
  With objExcel.Range("F" & (rowToStart + 4), "F" & (rowToStart + 5)).Borders(7)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 4), "F" & (rowToStart + 5)).Borders(8)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 4), "F" & (rowToStart + 5)).Borders(9)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With
  With objExcel.Range("F" & (rowToStart + 4), "F" & (rowToStart + 5)).Borders(10)
    .lineStyle = 1
    .weight = 3
    .colorIndex = -4105
  End With

  ' Date Completed
  objExcel.Range("G" & (rowToStart + 2), "I" & (rowToStart + 3)).Merge
  With objExcel.Range("G" & (rowToStart + 2))
    .Value = "Date Completed:"
    .Font.Size = 18
    .HorizontalAlignment = -4108
  End With
  objExcel.Range("G" & (rowToStart + 4), "I" & (rowToStart + 5)).Merge
  With objExcel.Range("G" & (rowToStart + 4), "I" & (rowToStart + 5)).Borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Notes Box
  With objExcel.Range("A" & (rowToStart + 7), "I" & (rowToStart + 22)).Borders(7)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("A" & (rowToStart + 7), "I" & (rowToStart + 22)).Borders(8)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("A" & (rowToStart + 7), "I" & (rowToStart + 22)).Borders(9)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("A" & (rowToStart + 7), "I" & (rowToStart + 22)).Borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Jobs
  objExcel.Range("A" & (rowToStart + 7), "C" & (rowToStart + 7)).Merge
  With objExcel.Range("A" & (rowToStart + 7), "C" & (rowToStart + 7)).Borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("A" & (rowToStart + 7), "C" & (rowToStart + 7)).Borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With
  with objExcel.Range("A" & (rowToStart + 7))
    .Value = "Jobs"
    .HorizontalAlignment = -4108
  End With
  iterator = 0
  For Each job in jobs
    With objExcel.Range("A" & (rowToStart + iterator + 8), "C" & (rowToStart + iterator + 8))
      .Merge
      With .Borders(9)
        .lineStyle = 5
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
    End With
    objExcel.Range("A" & (rowToStart + iterator + 8)).Value = job
    iterator = iterator + 1
  Next
  With objExcel.Range("A" & (rowToStart + iterator + 8), "C" & (rowToStart + iterator + 8)).Borders(8)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  ' Repairs / Notes
  objExcel.Range("D" & (rowToStart + 7), "F" & (rowToStart + 7)).Merge
  With objExcel.Range("D" & (rowToStart + 7), "F" & (rowToStart + 7)).Borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("D" & (rowToStart + 7))
    .Value = "Repairs / Notes"
    .HorizontalAlignment = -4108
  End With
  With objExcel.Range("F" & (rowToStart + 7), "F" & (rowToStart + 22)).Borders(10)
    .lineStyle = 1
    .weight = 2
    .colorIndex = -4105
  End With

  'Check if oil is due and put an X in the box if it is
  If workOrder.item("Oil") <> "" Then
    With objExcel.Range("D" & (rowToStart + 8), "F" & (rowToStart + 8))
      .Merge
      .Value = "Oil change due at " & workOrder.item("Oil") & " miles."
    End With
  Else
    With objExcel.Range("F" & (rowToStart + 4))
      .Font.Size = 24
      .HorizontalAlignment = -4108
      .VerticalAlignment = -4108
      .Value = "X"
    End With
  End If

  ' Parts Need to Order
  objExcel.Range("G" & (rowToStart + 7), "I" & (rowToStart + 7)).Merge
  With objExcel.Range("G" & (rowToStart + 7), "I" & (rowToStart + 7)).Borders(9)
    .lineStyle = -4118
    .weight = 2
    .colorIndex = -4105
  End With
  With objExcel.Range("G" & (rowToStart + 7))
    .Value = "Parts Need to Order"
    .HorizontalAlignment = -4108
  End With

  ' Tires
  With objExcel.Range("A" & (rowToStart + 18))
    .Value = "Tires"
    .HorizontalAlignment = -4108
  End With
  For i = 2 To 3
    With objExcel.Cells(rowToStart + 18, i)
      .Interior.colorIndex = 15
      With .Borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
    End With
  Next
  For i = 1 To 3
    With objExcel.Cells(rowToStart + 19, i)
      .Interior.colorIndex = 15
      With .Borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
    End With
  Next
  For i = 1 To 3
    With objExcel.Cells(rowToStart + 21, i)
      .Interior.colorIndex = 15
      With .Borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
    End With
  Next
  For i = 2 To 3
    With objExcel.Cells(rowToStart + 22, i)
      .Interior.colorIndex = 15
      With .Borders(7)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(8)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(9)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
      With .Borders(10)
        .lineStyle = 1
        .weight = 2
        .colorIndex = -4105
      End With
    End With
  Next
End Sub

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

On Error Resume Next
session.findById("wnd[0]").maximize

Dim truck()
Dim serviceInterval()
Dim jobs()
Dim userInput, i, x, verifyTrucks, trucksNotFound, trucksFound, result, workOrder, objExcel, objWorkbook, toggle
Set workOrder = CreateObject("Scripting.Dictionary")
toggle = "Top"

i = 0
' Ask the user which trucks they need
Do While True
    userInput = InputBox("What is the unit number of the next truck?" & vbCr & "Don't use a -" & vbCr & "If you're done, leave it blank.")
    If userInput = "" Then
        Exit Do
    End If
    Redim Preserve truck(i)
    truck(i) = userInput
    i = i + 1
Loop

' Ask the user for verification
Do While True
    i = 0
    verifyTrucks = ""
    For Each unit in truck
        verifyTrucks = verifyTrucks & vbCr & i + 1 & ": " & unit
        i = i + 1
    Next

    If MsgBox("Are all of these entered properly?" & verifyTrucks, vbYesNo) =  vbNo Then
        i = InputBox("What number do you need to change?" & "If you need to cancel the whole thing, leave blank." & verifyTrucks)
        If i = "" Then
            WScript.Quit
        End If
        i = i - 1
        userInput = InputBox("What is the new number for " & truck(i) & "?" & vbCr & "Leave blank to remove it.")
        If userInput <> "" Then
            truck(i) = userInput
        Else
            If i = UBound(truck) Then
                Redim Preserve truck(i - 1)
            End If
            If i <= UBound(truck)Then
                Do Until i => UBound(truck)
                    truck(i) = truck(i + 1)
                    i = i + 1
                Loop
                Redim Preserve truck(i - 1)
            End If
        End If
    Else
        Exit Do
    End If
Loop

For Each unit in truck
  On Error Resume Next
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unit
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
  
  Do
    If session.findById("wnd[0]/sbar").text = "No vehicles could be selected" Then
      trucksNotFound = trucksNotFound & vbCr & unit
      trucksFound = trucksFound & vbCr
      Exit Do
    Else
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
      Err.Clear

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
      convertServiceIntervalsToJobs()
      workOrder.add "Jobs", jobs
      readOrder()
      session.findById("wnd[0]/tbar[0]/btn[3]").press
    End If
  Loop While false
  On Error Goto 0
  addToSheet()
  workOrder.RemoveAll
  Redim jobs(0)
Next

If IsObject(objExcel) Then
  If PRODUCTION = true Then
    objWorkbook.PrintOut
  End If
  objWorkbook.Close False
  objExcel.workbooks.Close
  objExcel.Quit
End If

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press

If trucksNotFound <> "" Then
    MsgBox("These trucks were not found in SAP." & trucksNotFound)
End If
If trucksFound <> "" Then
    MsgBox("These trucks were created successfully." & trucksFound)
End If

Set objExcel = Nothing
Set objWorkbook = Nothing