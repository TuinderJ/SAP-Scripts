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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsalesrep"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_ERDAT1").text = "10/12/2024"
session.findById("wnd[0]/usr/ctxtS_WERKS1-LOW").text = "1401"
session.findById("wnd[0]/usr/ctxtS_WERKS1-LOW").setFocus
session.findById("wnd[0]/usr/ctxtS_WERKS1-LOW").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]/tbar[1]/btn[8]").press
