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
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER01"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZS21"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "100000"
session.findById("wnd[0]/usr/txt/DBM/ORDER_CREATION-LICPL").text = ""
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-LICPL_COUNTRY").text = ""
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VHVIN").text = ""
session.findById("wnd[0]/usr/chk/DBM/ORDER_CREATION-NO_VEHICLE").selected = true
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "7039"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").setFocus
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").setFocus
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-AC_AS_TYP").key = "923"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/ctxt/DBM/VBAK_COM-EMPGE").text = "703904"
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/ctxt/DBM/VBAK_COM-EMPGE").setFocus
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/ctxt/DBM/VBAK_COM-EMPGE").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "Shop Supplies"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = "TRANSFER TO SHOP SUPPLIES"
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").caretPosition = 25
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select

