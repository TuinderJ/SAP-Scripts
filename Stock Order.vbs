dim partsInStock()
partsMainFilePath = "C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\Parts Main.csv"

function storeSAPData()
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
  ' session.findById("wnd[0]").maximize

  session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
  session.findById("wnd[0]/tbar[0]/btn[0]").press

  session.findById("wnd[0]/usr/chkP_BIN").selected = true
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = ""
  session.findById("wnd[0]/usr/txtS_EMNFR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7039"
  session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "0001"
  session.findById("wnd[0]/usr/txtS_LGPBE-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MTART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = ""
  session.findById("wnd[0]/tbar[1]/btn[8]").press

  dim i
  on error resume next
  i = -1
  do while true
    i = i + 1
    redim preserve partsInStock(1,i)
    ' Part Number
    partsInStock(0,i) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"MATNR")
    ' Qty
    partsInStock(1,i) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"LABST")
    if err.number <> 0 then
      exit do
    end if
  loop
  err.clear
end function

function roundOrderQuantityToRoundingValue(orderQuantity, roundingValue)
  if roundingValue = 0 then
    roundOrderQuantityToRoundingValue = orderQuantity
    exit function
  end if  
  roundOrderQuantityToRoundingValue = cInt(orderQuantity / roundingValue) * roundingValue
end function

function changeCSV()
  'Create an Excel Object
  Set excel = createObject("Excel.Application")
  'Open the Rebill Pricing Excel File
  Set workbook = excel.workBooks.open(partsMainFilePath)

  Dim rowcount
  'Load the sheet
  Set sheet = workbook.worksheets("Parts Main")
  rowcount = sheet.Usedrange.Rows.Count

  dim ii
  For i = 2 To rowcount
    ii = 1
    do while true
      ii = ii + 1
      if ii > uBound(partsInStock,2) then
        sheet.cells(i, 6).value = 0
        exit do
      end if
      if sheet.cells(i, 3).value = partsInStock(0, ii) then
        sheet.cells(i, 6).value = partsInStock(1, ii)
        exit do
      end if
    loop
    dim orderQuantity, roundingValue
    orderQuantity = cInt(sheet.cells(i, 5).value) - cInt(sheet.cells(i, 6).value)
    if orderQuantity < 0 then
      orderQuantity = 0
    end if
    roundingValue = cInt(sheet.cells(i, 7).value)
    orderQuantity = roundOrderQuantityToRoundingValue(orderQuantity, roundingValue)
    sheet.cells(i, 2).value = orderQuantity
  Next

  'End of data, clear memory
  Set sheet = Nothing

  workbook.save
  workbook.close
  excel.workbooks.close
  excel.quit

  Set Workbook = Nothing
  Set excel = Nothing
End Function

function sendEmail()
  Set WshShell = WScript.CreateObject("WScript.Shell")

  Set outlook = CreateObject("Outlook.Application")
  Set email = outlook.CreateItem(0)

  With email
    .To = _
    "Dezell Hunter <hunterd1@RushEnterprises.com>; Malexy Rocha Olivas <RochaOlivasm@RushEnterprises.com>; Hunter Robinson <Robinsonh@RushEnterprises.com>"
    '.CC = ""
    '.BCC = ""
    .Subject = date & " Leasing Stock Order"
    ' .htmlBody = ""
    .Attachments.Add partsMainFilePath
    .Send
  End With

  Set outlook = Nothing
  Set email = Nothing
end function

storeSAPData()
changeCSV()
sendEmail()