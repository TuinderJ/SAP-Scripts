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
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "100000"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr").verticalScrollbar.position = 116
session.findById("wnd[1]/usr/lbl[22,30]").setFocus
session.findById("wnd[1]/usr/lbl[22,30]").caretPosition = 10
session.findById("wnd[1]").sendVKey 2
