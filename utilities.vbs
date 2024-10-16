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

sub goToTCode(tCode)
  session.findById("wnd[0]").maximize
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N" & tCode
  session.findById("wnd[0]/tbar[0]/btn[0]").press
end sub

sub goBackNScreens(n)
  for i = 0 to n - 1
    session.findById("wnd[0]").sendVKey 3
  next
end sub