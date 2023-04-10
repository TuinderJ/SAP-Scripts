Dim i, numberOfParts, numberOfOutputParts, objExcel, objWorkbook, partsWithNoBin(), outputPartsList()

Sub pullReport()
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
  session.findById("wnd[0]/usr/txtS_EMNFR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7013"
  session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "0001"
  session.findById("wnd[0]/usr/txtS_LGPBE-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MTART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = ""
  session.findById("wnd[0]/tbar[1]/btn[8]").press
End Sub

Sub findNoBinParts()
  numberOfParts = 0
  Do While True
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell numberOfParts,"MATNR"
    binLocation = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(numberOfParts,"LGPBE")
    If binLocation <> "" Then
      numberOfParts = numberOfParts - 1
      Exit Do
    End If
    Redim Preserve partsWithNoBin(1, numberOfParts)
    partNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(numberOfParts,"MATNR")
    partDescription = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(numberOfParts,"MAKTX")
    partsWithNoBin(0, numberOfParts) = partNumber
    partsWithNoBin(1, numberOfParts) = partDescription
    numberOfParts = numberOfParts + 1
  Loop
End Sub

Sub findHistoryOnNoBinParts()
  Dim scrollPosistion
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NMB51"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtMATNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "7013"
  session.findById("wnd[0]/usr/ctxtWERKS-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "0001"
  session.findById("wnd[0]/usr/ctxtLGORT-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtCHARG-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtLIFNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtLIFNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtKUNNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtKUNNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtBWART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtBWART-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtSOBKZ-HIGH").text = ""
  session.findById("wnd[0]/usr/txtMAT_KDAU-LOW").text = ""
  session.findById("wnd[0]/usr/txtMAT_KDAU-HIGH").text = ""
  session.findById("wnd[0]/usr/txtMAT_KDPO-LOW").text = ""
  session.findById("wnd[0]/usr/txtMAT_KDPO-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtSHKZG-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtSHKZG-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtUMWRK-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtUMWRK-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = date() - 14
  session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = date()
  session.findById("wnd[0]/usr/txtUSNAM-LOW").text = ""
  session.findById("wnd[0]/usr/txtUSNAM-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtVGART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtVGART-HIGH").text = ""
  session.findById("wnd[0]/usr/txtXBLNR-LOW").text = ""
  session.findById("wnd[0]/usr/txtXBLNR-HIGH").text = ""
  session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press

  scrollPosistion = 0
  On Error Resume Next
  For i = 0 To numberOfParts Step 1
    If Err.Number <> 0 Then
      scrollPosistion = scrollPosistion + 1
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = scrollPosistion
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,7]").text = partsWithNoBin(0, i)
    Else
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & i & "]").text = partsWithNoBin(0, i)
      session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & (i + 1) & "]").setFocus
    End If
  Next
  Err.Clear
  session.findById("wnd[1]/tbar[0]/btn[8]").press
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  session.findById("wnd[0]/tbar[1]/btn[48]").press
  i = 0
  numberOfOutputParts = 0
  Do While True
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell i, "MATNR"
    If err.number <> 0 Then
      Exit Do
    End If
    partNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "MATNR")
    partDescription = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "MAKTX")
    movementType = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "BWART")
    reference = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "XBLNR")
    qty = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "ERFMG")
    If Right(qty, 1) = "-" Then
      qty = "-" & Left(qty, Len(qty) - 1)
    End If
    nextPartNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i + 1, "MATNR")
    If partNumber <> nextPartNumber Then
      If movementType = "101" Then
        Redim Preserve outputPartsList(4, numberOfOutputParts)
        outputPartsList(0, numberOfOutputParts) = partNumber
        outputPartsList(1, numberOfOutputParts) = partDescription
        outputPartsList(2, numberOfOutputParts) = qty
        outputPartsList(3, numberOfOutputParts) = reference
        numberOfOutputParts = numberOfOutputParts + 1
      End If
    Else
      checkMultipleLines()
    End If
    i = i + 1
  Loop
  numberOfOutputParts = numberOfOutputParts - 1
End Sub

Sub checkMultipleLines()
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell i, "MATNR"
  partNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "MATNR")
  partDescription = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "MAKTX")
  movementType = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "BWART")
  reference = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "XBLNR")
  qty = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i, "ERFMG")
  If Right(qty, 1) = "-" Then
    qty = "-" & Left(qty, Len(qty) - 1)
  End If
  nextPartNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(i + 1, "MATNR")
  Redim Preserve outputPartsList(4, numberOfOutputParts)
  outputPartsList(0, numberOfOutputParts) = partNumber
  outputPartsList(1, numberOfOutputParts) = partDescription
  outputPartsList(2, numberOfOutputParts) = qty
  outputPartsList(3, numberOfOutputParts) = reference
  If movementType <> "101" Then
    session.findById("wnd[0]").sendVKey 2
    outputPartsList(4, numberOfOutputParts) = Right(Left(session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/txtGOITEM-SGTXT[40,0]").Text, 18), 8)
    session.findById("wnd[0]/tbar[0]/btn[3]").press
  End If
  numberOfOutputParts = numberOfOutputParts + 1
  i = i + 1
  If partNumber = nextPartNumber Then
    checkMultipleLines()
  End If
End Sub

Sub makeSpreadsheet()
  Set objExcel = CreateObject("Excel.Application")
  Set objWorkbook = objExcel.Workbooks.Add()
  objExcel.visible = True
  objExcel.WindowState = -4137

  objExcel.Range("A1").Value = "Part Number"
  objExcel.Range("B1").Value = "Part Description"
  objExcel.Range("C1").Value = "Qty"
  objExcel.Range("D1").Value = "Reference"
  objExcel.Range("E1").Value = "RO"
  For i = 0 To numberOfOutputParts
    With objExcel.Range("A" & (i + 2), "E" & (i + 2)).Borders(7)
      .lineStyle = 1
      .weight = 2
      .colorIndex = -4105
    End With
    With objExcel.Range("A" & (i + 2), "E" & (i + 2)).Borders(8)
      .lineStyle = 1
      .weight = 2
      .colorIndex = -4105
    End With
    With objExcel.Range("A" & (i + 2), "E" & (i + 2)).Borders(9)
      .lineStyle = 1
      .weight = 2
      .colorIndex = -4105
    End With
    With objExcel.Range("A" & (i + 2), "E" & (i + 2)).Borders(10)
      .lineStyle = 1
      .weight = 2
      .colorIndex = -4105
    End With
    objExcel.Range("A" & (i + 2)).Value = outputPartsList(0, i)
    objExcel.Range("B" & (i + 2)).Value = outputPartsList(1, i)
    objExcel.Range("C" & (i + 2)).Value = outputPartsList(2, i)
    objExcel.Range("D" & (i + 2)).Value = outputPartsList(3, i)
    objExcel.Range("E" & (i + 2)).Value = outputPartsList(4, i)
  Next
  For i = 1 To 5
    objWorkbook.Sheets("Sheet1").columns(i).AutoFit()
  Next
End Sub

Sub saveSpreadsheetAndSendEmail()
  Set WshShell = WScript.CreateObject("WScript.Shell")
  strDesktop = WshShell.SpecialFolders("Desktop")
  strNewExcelFilePath = strDesktop & "\No Bill Parts 14 Days.xlsx"
  objWorkbook.SaveAs(strNewExcelFilePath)
  objWorkbook.Close
  objExcel.workbooks.Close
  objExcel.Quit
  Set objWorkbook = Nothing
  Set objExcel = Nothing

  Set objOutlook = CreateObject("Outlook.Application")
  Set objEmail = objOutlook.CreateItem(0)

  With objEmail
    .To = _
    "tuinderj@rushenterprises.com; " & _
    "Arreya@rushenterprises.com; " & _
    "ramirezm4@rushenterprises.com"
    '.CC = ""
    '.BCC = ""
    .Subject = "Parts Not Billed from past 14 days"
    ' .htmlBody = ""
    .Attachments.Add strNewExcelFilePath
    .Send
  End With

  Set objOutlook = Nothing
  Set objEmail = Nothing
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

' On Error Resume Next
session.findById("wnd[0]").maximize

pullReport()
findNoBinParts()
findHistoryOnNoBinParts()
makeSpreadsheet()
If MsgBox("Review the spreadsheet." & vbCr & "If you want to proceed with the email, press OK." & vbCr & "If you don't want to continue, press Cancel.", vbOkCancel, "Pause for Manual Review") = vbCancel Then
  WScript.Quit
End If
saveSpreadsheetAndSendEmail()