Dim production, debug, readOnly
production = false
debug = false
readOnly = false

Dim ro, vin, customer, orderType, advisor, mileage, hours, orderStatus, accountAssignmentCategory, po, headerText, jobs(), laborLines(), labor(), parts(), manualParts()

Function findJobNumber(jobNumberInput)
  For i = 0 To UBound(jobs,2)
    If cInt(jobs(3, i)) = cInt(jobNumberInput) Then
      findJobNumber = i + 1
    End If
  Next
End Function

Sub readHeader()
  'Go to the RO
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = ro
  session.findById("wnd[0]").sendVKey 0

  vin = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/txt/DBM/VEHORDCOM-VHVIN").text
  customer = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/ctxt/DBM/VBAK_COM-PARTNER").text
  orderType = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/ctxt/DBM/VBAK_COM-AUFART").text
  advisor = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/ctxt/DBM/VBAK_COM-PERNR").text
  mileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text
  hours = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text
  orderStatus = trim(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-ZZORD_STATUS").text)
  accountAssignmentCategory = trim(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").text)
  po = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-BSTNK").text

  'Open header text
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
  headerText = session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text
  session.findById("wnd[1]/tbar[0]/btn[8]").press



  If debug = true Then
    msgBox(_
    "Header" & vbCr & _
    "Vin:" & vbCr & vin & vbCr & vbCr & _
    "Customer:" & vbCr & customer & vbCr & vbCr & _
    "RO:" & vbCr & ro & vbCr & vbCr & _
    "Order Type:" & vbCr & orderType & vbCr & vbCr & _
    "Advisor Number:" & vbCr & advisor & vbCr & vbCr & _
    "Mileage:" & vbCr & mileage & vbCr & vbCr & _
    "Hours:" & vbCr & hours & vbCr & vbCr & _
    "Order Status:" & vbCr & orderStatus & vbCr & vbCr & _
    "Account Assignment Category:" & vbCr & accountAssignmentCategory & vbCr & vbCr & _
    "Po:" & vbCr & po & vbCr & vbCr & _
    "Header Text:" & vbCr & headerText _
    )
  End If
End Sub

Sub readJobs()
  'Select Jobs tab  
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  
  'Set up jobs selectors and itterator
  Dim a, b, c, d, i
  a = 0
  b = 0
  c = 0
  d = 0
  i = -1

  Dim continue
  
  On Error Resume Next
  Do Until d & c & b & a > 1000
    'Increment jobs selectors
    a = a + 1
    If a = 10 Then
      b = b + 1
      a = 0
    End If
    If b = 10 Then
      c = c + 1
      b = 0
    End If
    If c = 10 Then
      d = d + 1
      c = 0
    End If
    
    'Double click on job
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA4:/DBM/SAPLORDER_UI:2053/subSUBSCREEN_2053:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2323/cntlTREE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "J00" & d & c & b & a,"1"
    'Error check for if there was no job to select
		If Err.Number = 0 Then
      If "CORES" <> session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text Then
      i = i + 1
        Redim Preserve jobs(3, i)
        'Read description
        jobs(0, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text
        'Click into the story
        session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/btnJOB_LONG_TEXT").press
        'Read workshop text
        jobs(1, i) = session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text
        'Read invoice text
        session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "0002","COLUMN1"
        session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "0002EN","COLUMN1"
        jobs(2, i) = session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text
        'Close texts
        session.findById("wnd[1]/tbar[0]/btn[12]").press
        'Old Job Number
        jobs(3, i) = d & c & b & a
      End If
    End If
      Err.Clear
  Loop



  If debug = true Then
    For ii = 0 To UBound(jobs, 2)
      msgBox(_
      "Description:" & vbCr & jobs(0, ii) & vbCr & vbCr & _
      "Workshop Text:" & vbCr & jobs(1, ii) & vbCr & vbCr & _
      "Invoice Text:" & vbCr & jobs(2, ii))
    Next
  End If
End Sub

Sub readLaborLines()
  'Select Item tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "JOBS"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "ITCAT"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").pressToolbarButton "&SORT_ASC"
  Redim laborLines(2, 0)
  Dim i, row, temp
  i = 0
  row = 0

  On Error Resume Next
  Do Until Err.Number <> 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row + 1,"ITCAT"
    'See if current line is a lobor value
    If "P001" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITCAT") Then
      Redim Preserve laborLines(2, i)
      'Get job
      laborLines(0, i) = findJobNumber(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"JOBS"))
      'Get SRT key
      laborLines(1, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITOBJID")
      'Get description
      laborLines(2, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"DESCR1")

      i = i + 1
    End If
    row = row + 1
  Loop
  Err.Clear

  If debug = true Then
    For ii = 0 To UBound(laborLines, 2)
      msgBox("Labor line: " & ii & vbCr & vbCr & "Job: " & vbCr & laborLines(0, ii) & vbCr & vbCr & "SRT: " & vbCr & laborLines(1, ii) & vbCr & vbCr & "Description: " & vbCr & laborLines(2, ii))
    Next
  End If
End Sub

Sub readLabor()
  On Error Resume Next
  Redim labor(3, 0)
  'Press Labor button
  session.findById("wnd[0]/tbar[1]/btn[41]").press
  If Err.number <> 0 Then Exit Sub
  Dim i, row
  i = 0
  row = 0

  Do Until Err.Number <> 0
    session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").setCurrentCell row + 1,"JOB"
    If session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").getCellValue(row,"TECH") <> "" Then
      Redim Preserve labor(3, i)
      'Get job number
      labor(0, i) = findJobNumber(session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").getCellValue(row,"JOB"))
      'Get tech number
      labor(1, i) = session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").getCellValue(row,"TECH")
      'Get date
      labor(2, i) = session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").getCellValue(row,"STARTDATE")
      'Get total hours
      labor(3, i) = session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").getCellValue(row,"THOURS")
      i = i + 1
    End If
    row = row + 1
  Loop
  Err.Clear
  session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "EXIT"





  If debug = true Then
    For ii = 0 To UBound(labor, 2)
      msgBox(_
      "Labor: " & vbCr & ii & vbCr & vbCr & _
      "Job:" & vbCr & labor(0, ii) & vbCr & vbCr & _
      "Tech Number:" & vbCr & labor(1, ii) & vbCr & vbCr & _
      "Date:" & vbCr & labor(2, ii) & vbCr & vbCr & _
      "Hours:" & vbCr & labor(3, ii))
    Next
  End If
End Sub

Sub readParts()
  'Select Item tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "JOBS"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "ITCAT"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").pressToolbarButton "&SORT_ASC"
  Redim parts(4, 0)
  Dim i, row
  i = 0
  row = 0

  On Error Resume Next
  'Select next item
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
    Do Until Err.Number <> 0
    'See if current line is a part
    If "P002" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITCAT") Then
      Redim Preserve parts(4, i)
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell
      'Get job
      parts(0, i) = findJobNumber(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"JOBS"))
      'Get quantity
      parts(1, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ZMENG")
      'Get part number
      parts(2, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITOBJID")
      'Get description
      parts(3, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"DESCR1")
      'Get manual price
      parts(4, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3311/txt/DBM/S_POS-KBETM").text

      i = i + 1
    End If
    row = row + 1
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
  Loop
  Err.Clear



  If debug = true Then
    For ii = 0 To UBound(parts, 2)
      msgBox(_
      "Part:" & vbCr & ii & vbCr & vbCr & _
      "Job:" & vbCr & parts(0, ii) & vbCr & vbCr & _
      "Quantity:" & vbCr & parts(1, ii) & vbCr & vbCr & _
      "Part Number:" & vbCr & parts(2, ii) & vbCr & vbCr & _
      "Description" & vbCr & parts(3, ii) & vbCr & vbCr & _
      "Manual Price" & vbCr & parts(4, ii))
    Next
  End If
End Sub

Sub readManualParts()
  'Select Item tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "JOBS"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").selectColumn "ITCAT"
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").pressToolbarButton "&SORT_ASC"
  Redim manualParts(5, 0)
  Dim i, row
  i = 0
  row = 0

  On Error Resume Next
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
    Do Until Err.Number <> 0
    'Select next item

    'See if current line is a part
    If "P009" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITCAT") Then
      Redim Preserve manualParts(5, i)
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").doubleClickCurrentCell
      'Get job
      manualParts(0, i) = findJobNumber(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"JOBS"))
      'Get quantity
      manualParts(1, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ZMENG")
      'Get part number
      manualParts(2, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITOBJID")
      'Get description
      manualParts(3, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"DESCR1")
      'Get manual price
      manualParts(4, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-KBETM").text
      'Get purchase price
      manualParts(5, i) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-VERPR").text

      i = i + 1
    End If
    row = row + 1
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
  Loop
  Err.Clear



  If debug = true Then
    For ii = 0 To UBound(manualParts, 2)
      msgBox(_
      "Manual Part:" & vbCr & ii & vbCr & vbCr & _
      "Job" & vbCr & manualParts(0, ii) & vbCr & vbCr & _
      "Quantity" & vbCr & manualParts(1, ii) & vbCr & vbCr & _
      "Part Number" & vbCr & manualParts(2, ii) & vbCr & vbCr & _
      "Description" & vbCr & manualParts(3, ii) & vbCr & vbCr & _
      "Manual Price" & vbCr & manualParts(4, ii) & vbCr & vbCr & _
      "Purchase Price" & vbCr & manualParts(5, ii))
    Next
  End If
End Sub

Sub deleteOrderIfThereAreNoPos()
  'Select Item tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  Dim row, roContainsPo
  row = 0

  On Error Resume Next
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
  Do Until Err.Number <> 0
    'See if current line is a lobor value
    If "P010" = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").getCellValue(row,"ITCAT") Then
      roContainsPo = true
    End If
    row = row + 1
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2300/cntlITEM_ALV_CUSTOM_CONTAINER/shellcont/shell").setCurrentCell row,"ITCAT"
  Loop
  Err.Clear

  If Not roContainsPo Then
    'Open RO
    session.findById("wnd[0]/tbar[1]/btn[13]").press
    If production = true Then
      'Cancel all parts goods movement
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectAll
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressToolbarButton "ITEM_CANCEL"
      If labor(0, 0) <> "" Then
        'Remove all labor
        session.findById("wnd[0]/tbar[1]/btn[41]").press
        session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").selectAll
        session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "DEL"
        session.findById("wnd[2]/usr/btnBUTTON_1").press
        session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "EXIT"
      End If
    End If
    'Delete
    session.findById("wnd[0]").sendVKey 14
    If production = true Then
      session.findById("wnd[1]/usr/btnBUTTON_1").press
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      Err.Clear
    Else
      session.findById("wnd[1]/usr/btnBUTTON_2").press
    End If
  End If
End Sub

Sub makeNewRO()
  'Go to Order Processing and input vehicle data
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  'Reset fields
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press
  'Input VIN
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-VHVIN").text = VIN
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON05").press
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = customer
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/txt/DBM/ORDER_SEARCH-BSTNK").text = po
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/cmb/DBM/ORDER_SEARCH-AUFART").key = orderType
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PERNR").text = advisor
  session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON03").press

  'Error handling
  On Error Resume Next
  Dim errorMessage
  errorMessage = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0,"T_MSG")
  If errorMessage <> "" Then
    Do Until errorMessage = ""
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/ctxt/DBM/ORDER_SEARCH-PARTNER").text = InputBox(errorMessage + vbCr + "Please select a different customer number","Customer Number")
      CheckCustomer()
      session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON03").press
      errorMessage = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0,"T_MSG")
      If Err.Number <> 0 Then
        errorMessage = ""
      End If
    Loop
  End If

  session.findById("wnd[1]/tbar[0]/btn[0]").press
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  Err.Clear

  '------------------Header------------------
  'Fill out header
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = mileage
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = hours
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-ZZORD_STATUS").key = orderStatus
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = accountAssignmentCategory
  Err.clear

  'Header text
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
  session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = headerText
  session.findById("wnd[1]/tbar[0]/btn[8]").press
  
  '------------------Jobs------------------
  '0 = Job Title
  '1 = Workshop Text
  '2 = Invoice Text

  'Select Jobs tab  
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select

  'Fill out jobs
  For i = 0 To UBound(jobs, 2)
    'Job title
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = jobs(0, i)
    session.findById("wnd[0]").sendVKey 0
    'Texts
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/btnJOB_LONG_TEXT").press
    session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = jobs(1, i)
    session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "0002","COLUMN1"
    session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = jobs(2, i)
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    'Clear for next job
    session.findById("wnd[0]").sendVKey 5
  Next
  '------------------Labor Lines------------------
  '0 = Job
  '1 = SRT Key
  '2 = Description

  If laborLines(0, 0) <> "" Then

    'Select Item tab
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select

    'Fill out labor lines
    For i = 0 To UBound(laborLines, 2)
      'SRT Key
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = laborLines(1, i)
      session.findById("wnd[0]").sendVKey 0
      'Job
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = laborLines(0, i)
      'Description
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/txt/DBM/S_POS-DESCR1").text = laborLines(2, i)
      session.findById("wnd[0]").sendVKey 0
    Next

  End If

  'Save/release/save (if there is any clocked labor)
  session.findById("wnd[0]/tbar[0]/btn[11]").press
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  Err.Clear
  If labor(0, 0) <> "" Then
    session.findById("wnd[0]/tbar[1]/btn[37]").press
    session.findById("wnd[0]/tbar[0]/btn[11]").press
  End If

  '------------------Labor------------------
  '0 = Job
  '1 = Tech
  '2 = Date
  '3 = Hours

  If labor(0, 0) <> "" Then
    'Press Labor button
    session.findById("wnd[0]/tbar[1]/btn[41]").press

    For i = 0 To UBound(labor, 2)
      session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "ADD"
      Dim row
      row = 0
      Do Until Err.Number <> 0
        session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").setCurrentCell row,"JOB"
        row = row + 1
      Loop
      row = row - 3
      Err.clear
      session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").modifyCell row,"JOB",labor(0, i)
      session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").modifyCell row,"TECH",labor(1, i)
      session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").modifyCell row,"STARTDATE",labor(2, i)
      session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").modifyCell row,"THOURS",labor(3, i)
      session.findById("wnd[1]").sendVKey 0
    Next
      
    session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").selectAll
    session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "TIME"
    session.findById("wnd[1]/usr/cntlCC_TIME_NEW/shellcont/shell").pressToolbarButton "EXIT"

  End If
  Err.Clear

  '------------------Parts------------------
  '0 = Job
  '1 = Quantity
  '2 = Part Number
  '3 = Description
  '4 = Manual Price

  If parts(0, 0) <> "" Then
    'Select Parts tab
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select

    For i = 0 To UBound(parts, 2)
      'Part Number
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = parts(2, i)
      session.findById("wnd[0]").sendVKey 0
      'Quantity
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = parts(1, i)
      'Description
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-DESCR1").text = parts(3, i)
      'Manual Price
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-KBETM").text = parts(4, i)
      'Job
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-JOBS").text = parts(0, i)
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[1]/tbar[0]/btn[0]").press
      'Core
      session.findById("wnd[1]/usr/btnCORE1").press
    Next
  End If
  Err.Clear

  '------------------Manual Parts------------------
  '0 = Job
  '1 = Quantity
  '2 = Part Number
  '3 = Description
  '4 = Manual Price
  '5 = Purchase Price

  If manualParts(0, 0) <> "" Then
    'Select Item tab
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
    'Select manual part
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3311/cmb/DBM/S_POS-ITCAT").key = "P009"

    For i = 0 To UBound(manualParts, 2)
      'Part Number
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-ITOBJID").text = manualParts(2, i)
      'Description
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-DESCR1").text = manualParts(3, i)
      'Quantity
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-ZMENG").text = manualParts(1, i)
      'Job
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/ctxt/DBM/S_POS-JOBS").text = manualParts(0, i)
      'Purchase Price
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-VERPR").text = manualParts(5, i)
      'Manual Price
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/txt/DBM/S_POS-KBETM").text = manualParts(4, i)
      'Material Group
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/ctxt/DBM/S_POS-MATKL").text = "2000"
      'Price Reference Material
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3314/ctxt/DBM/S_POS-MATNR18").text = "MANPARTTX"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/tbar[0]/btn[11]").press
    Next
  End If

  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01").select
End Sub

'-----------------------------------------------------------------------------------------------------
'BEGIN
'-----------------------------------------------------------------------------------------------------
' User input for locating the RO
Do Until Len(ro) = 8
If ro = "" Then
   ro = inputBox("What is the RO number thgat you would like to recreate?", "RO Number")
Else
	ro = InputBox("Please type a valid RO number.", "RO Number",Inv)
End If
If ro = "" Then
	WScript.Quit
End If
ro = Trim(ro)
Loop
' ro = 39257977


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
session.findById("wnd[0]").maximize

readHeader()
readJobs()
readLaborLines()
readLabor()
readParts()
readManualParts()
deleteOrderIfThereAreNoPos()
If readOnly = false Then
  makeNewRO()
End If