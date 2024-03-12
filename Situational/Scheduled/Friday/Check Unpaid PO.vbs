pullReport()
readSpreadsheet()
sendEmail()

dim total, total7008, total7013, total7020, total7039, strDesktop

function pullReport()
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
  ' Pull Report
  session.findById("wnd[0]").maximize
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NME2N"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "GDSRCPT"
  session.findById("wnd[0]/usr/ctxtEN_EBELN-LOW").text = "4500000000"
  session.findById("wnd[0]/usr/ctxtEN_EBELN-HIGH").text = "5000000000"
  session.findById("wnd[0]/usr/ctxtEN_EKORG-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtEN_EKORG-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtEN_MFRPN-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtEN_MFRPN-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtSELPA-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_BSART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_BSART-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_EKGRP-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_WERKS-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_PSTYP-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_PSTYP-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_KNTTP-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_EINDT-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtP_GULDT").text = ""
  session.findById("wnd[0]/usr/ctxtP_RWEIT").text = ""
  session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_LIFNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_RESWK-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_RESWK-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATKL-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = date - 90
  session.findById("wnd[0]/usr/txtS_EAN11-LOW").text = ""
  session.findById("wnd[0]/usr/txtS_EAN11-HIGH").text = ""
  session.findById("wnd[0]/usr/txtS_IDNLF-LOW").text = ""
  session.findById("wnd[0]/usr/txtS_IDNLF-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_LTSNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_LTSNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_AKTNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_AKTNR-HIGH").text = ""
  session.findById("wnd[0]/usr/ctxtS_SAISO-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_SAISO-HIGH").text = ""
  session.findById("wnd[0]/usr/txtS_SAISJ-LOW").text = ""
  session.findById("wnd[0]/usr/txtS_SAISJ-HIGH").text = ""
  session.findById("wnd[0]/usr/txtP_TXZ01").text = ""
  session.findById("wnd[0]/usr/txtP_NAME1").text = ""
  session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "7008"
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "7013"
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "7020"
  session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "7039"
  session.findById("wnd[1]/tbar[0]/btn[8]").press
  session.findById("wnd[0]/tbar[1]/btn[8]").press

  ' Export to spreadsheet
  session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
  session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\"
  session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Unpaid PO's.xlsx"
  session.findById("wnd[1]/tbar[0]/btn[11]").press
end function

function readSpreadsheet()
  wScript.sleep 1500
  set objExcel = getObject(,"Excel.Application")
  wScript.sleep 1500
  set objWorkbook = objExcel.workbooks("Unpaid PO's.xlsx")

  ' Read totals here
  total = objExcel.cells(2, 13).value
  dim row
  row = 3
  do while true
    if objExcel.cells(row, 1).value = "" then
      exit do
    end if
    if objExcel.cells(row, 2).value = "" then
      select case objExcel.cells(row, 1).value
        case "7008"
          total7008 = objExcel.cells(row, 13).value
        case "7013"
          total7013 = objExcel.cells(row, 13).value
        case "7020"
          total7020 = objExcel.cells(row, 13).value
        case "7039"
          total7039 = objExcel.cells(row, 13).value
      end select
    end if
    row = row + 1
  loop

  objWorkbook.close
  set wshShell = nothing
  set objExcel = nothing
  set objWorkbook = nothing
end function

function sendEmail()
  dim htmlOutput, strNewExcelFilePath
  
  set objOutlook = createObject("Outlook.Application")
  set objEmail = objOutlook.createItem(0)
  
  set wshShell = wScript.createObject("WScript.Shell")
  strDesktop = wshShell.specialFolders("Desktop")

  strNewExcelFilePath = strDesktop & "\Unpaid PO's.xlsx"

  htmlOutput = _
    "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCr & _
    "<html>" & vbCr & _
      "<head>" & vbCr & _
        "<meta name=""Generator"" content=""MS Exchange Server version 16.0.13801.20804""/>" & vbCr & _
      "</head>" & vbCr & _
      "<body>" & vbCr & _
        "<p>" & vbCr & _
          "<font face=""Calibri"">" & vbCr & _
            "Current unpaid PO total cost." & vbCr & _
            "<br />" & vbCr & _
            "Region:" & vbCr & _
            "$" & formatNumber(total, 2) & vbCr & _
            "<br />" & vbCr & _
            "<br />" & vbCr & _
            "Branch: 7008" & vbCr & _
            "<br />" & vbCr & _
            "Total: $" & formatNumber(total7008, 2) & vbCr & _
            "<br />" & vbCr & _
            "<br />" & vbCr & _
            "Branch: 7013" & vbCr & _
            "<br />" & vbCr & _
            "Total: $" & formatNumber(total7013, 2) & vbCr & _
            "<br />" & vbCr & _
            "<br />" & vbCr & _
            "Branch: 7020" & vbCr & _
            "<br />" & vbCr & _
            "Total: $" & formatNumber(total7020, 2) & vbCr & _
            "<br />" & vbCr & _
            "<br />" & vbCr & _
            "Branch: 7039" & vbCr & _
            "<br />" & vbCr & _
            "Total: $" & formatNumber(total7039, 2) & vbCr & _
            "<br />" & vbCr & _
            "</font>" & vbCr & _
            "</p>" & vbCr & _
            "<br />" & vbCr & _
            "<br />" & vbCr & _
            "<p>" & vbCr & _
            "<font face=""Calibri"">" & vbCr & _
            "<b>Joshua Tuinder" & vbCr & _
            "<br />" & vbCr & _
            "Rush Truck Leasing" & vbCr & _
            "<br />" & vbCr & _
            "Mountain Region Inventory Control Supervisor</b>" & vbCr & _
            "<br />" & vbCr & _
            "379 W 66th Way" & vbCr & _
            "<br />" & vbCr & _
            "Denver, CO" & vbCr & _
            "<br />" & vbCr & _
            "O: (720) 292-5808" & vbCr & _
            "<br />" & vbCr & _
            "C: (720) 413-1681" & vbCr & _
          "</font>" & vbCr & _
          "<br />" & vbCr & _
          "<img src=""C:\Users\tuinderj\OneDrive - Rush Enterprises\Pictures\leasing-logo.png"" />" & vbCr & _
        "</p>" & vbCr & _
      "</body>" & vbCr & _
    "</html>"

  with objEmail
    .to = _
        "elliottr@rushenterprises.com; " &_
        "pilottem@rushenterprises.com; " &_
        "arreya@rushenterprises.com; " &_
        "martinezj15@rushenterprises.com; " &_
        "rings@rushenterprises.com; " &_
        "JacksonC1@RushEnterprises.com; " &_
        "Arndtb@RushEnterprises.com; " &_
        "DavilaV@RushEnterprises.com; " &_
        "TrevizoR@RushEnterprises.com"
    .cc = ""
    .bcc = _
        "tuinderj@rushenterprises.com"
    .subject = "Unpaid PO's " & Month(Date) & "/" & Day(Date)
    .htmlBody = htmlOutput
    .attachments.Add strNewExcelFilePath
    .send
  end with

  set objOutlook = nothing
  set objEmail = nothing
end function