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

Dim i
Dim part()

session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7039"
session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "0002"
session.findById("wnd[0]/tbar[1]/btn[8]").press

i = 0
Do While Err.Number = 0
    ReDim Preserve part(1, i)
    part(0, i) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"MATNR")
    part(1, i) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"LABST")
    i = i + 1
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = CStr(i)
Loop

session.findById("wnd[0]/tbar[0]/okcd").text = "/NME21N"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").key = "UB"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = "7039"

On Error Goto 0

i = 0
Do While i <= UBound(part, 2)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[3," & CStr(i) & "]").text = part(0, i)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[5," & CStr(i) & "]").text = part(1, i)
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[9," & CStr(i) & "]").text = "1401"
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-LGOBE[10," & CStr(i) & "]").text = "0002"
    i = i + 1
Loop

session.findById("wnd[0]").sendVKey 0
' session.findById("wnd[0]").sendVKey 11