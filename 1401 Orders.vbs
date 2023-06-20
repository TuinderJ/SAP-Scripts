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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nZSOLIST"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtS_KUNNR-LOW").text = "7039"
session.findById("wnd[0]/usr/ctxtS_VKORG-LOW").text = "1000"
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "1401"
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "VBELN"
session.findById("wnd[0]/tbar[1]/btn[28]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "NAME1",3
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "NAME2",3
