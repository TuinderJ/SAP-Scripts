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
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "LF14000NNFLG:FG"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "FF5825NNFLG:FG"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "FS19764FLG:FG"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").pressEnter
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "SR2000092:PB"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").pressEnter
