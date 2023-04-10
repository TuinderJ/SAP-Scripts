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

session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "7008"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "7013"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "7020"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "7039"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "0002"
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "EXCOST"
session.findById("wnd[0]/tbar[1]/btn[30]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WERKS"
session.findById("wnd[0]/tbar[1]/btn[42]").press

session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Cores & No Move.XLSX"
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3

WScript.Sleep 2500
Set objExcel = GetObject(,"Excel.Application")
Set objWorkbook = objExcel.Workbooks("Cores & No Move.XLSX")

objExcel.Cells(1, 2).WrapText = False
objExcel.Cells.EntireColumn.AutoFit
objExcel.ActiveSheet.Name = "Cores"

Set objExcel = Nothing
Set objWorkbook = Nothing