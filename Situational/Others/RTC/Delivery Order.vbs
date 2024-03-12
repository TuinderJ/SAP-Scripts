Dim advisorNumber, po

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

advisorNumber = "73363"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

po = inputBox("What is the PO that you would like to use?", "PO")
If po = "" Then
   WScript.quit
End If

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
On Error Resume Next

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER01"
session.findById("wnd[0]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP21"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "7039"
session.findById("wnd[0]/usr/chk/DBM/ORDER_CREATION-NO_VEHICLE").selected = True
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = advisorNumber
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickItem "ZSTO","COLUMN1"
session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = po
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "D"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text = po