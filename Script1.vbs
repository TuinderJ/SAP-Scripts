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
session.findById("wnd[0]/usr/tabsMAIN/tabpVEHDETAIL/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:SAPLZZGC001_01:7100/tabsDATAENTRY/tabpDATAENTRY_FC1/ssubDATAENTRY_SCA:SAPLZZGC001_01:9100/ctxtVLCACTDATA_ITEM_S-VHCEX").setFocus
session.findById("wnd[0]/usr/tabsMAIN/tabpVEHDETAIL/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/ssubDETAIL_SUBSCR:SAPLZZGC001_01:7100/tabsDATAENTRY/tabpDATAENTRY_FC1/ssubDATAENTRY_SCA:SAPLZZGC001_01:9100/ctxtVLCACTDATA_ITEM_S-VHCEX").caretPosition = 3
