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
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00068"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00071"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00072"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00096"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00069"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00067"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode "F00070"
session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F00052"
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
