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
session.findById("wnd[0]/tbar[0]/okcd").text = "/NZPMDUE"
session.findById("wnd[0]/tbar[0]/btn[0]").press
If Month(Now) = 12 Then
   session.findById("wnd[0]/usr/ctxtP_DNEXT").text = "1/1/" & Year(Now) + 1
Else
   session.findById("wnd[0]/usr/ctxtP_DNEXT").text = Month(Now) + 1 & "/1/" & Year(Now)
End If
session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7039"
session.findById("wnd[0]/usr/chkP_OILDR").selected = true
session.findById("wnd[0]/usr/chkP_PMONLY").selected = true
session.findById("wnd[0]").sendVKey 8