dim workbookName
workbookName = "Weekly Inventory.XLSX"

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

' Pull Bin Location
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
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select

' Save the file
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = workbookName
session.findById("wnd[1]/tbar[0]/btn[11]").press

' Run the Excel Macro
WScript.sleep 3000
set excel = getObject(,"Excel.Application")
set workbook = excel.workbooks(workbookName)
Workbook.Application.Run "PERSONAL.XLSB!FormatInventory"

set workbook = nothing
set excel = nothing