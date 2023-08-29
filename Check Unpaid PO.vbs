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
' branch = inputBox("What branch would you like to check for?", "Branch", 7039)
' if branch = "" then
'   Wscript.Quit
' end if

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/NME2N"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtSELPA-LOW").text = "GDSRCPT"
session.findById("wnd[0]/usr/ctxtEN_EBELN-LOW").text = "4500000000"
session.findById("wnd[0]/usr/ctxtEN_EBELN-HIGH").text = "5000000000"
session.findById("wnd[0]/usr/ctxtEN_EKORG-LOW").text = ""
session.findById("wnd[0]/usr/ctxtEN_EKORG-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtEN_MFRPN-LOW").text = ""
session.findById("wnd[0]/usr/ctxtEN_MFRPN-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtSELPA-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_BSART-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_BSART-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_EKGRP-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_EKGRP-HIGH").text = ""
' session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = branch
session.findById("wnd[0]/usr/ctxtS_WERKS-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_PSTYP-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_PSTYP-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_KNTTP-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_KNTTP-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_EINDT-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_EINDT-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtP_GULDT").text = ""
session.findById("wnd[0]/usr/ctxtP_RWEIT").text = ""
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_LIFNR-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_RESWK-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_RESWK-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_MATKL-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_BEDAT-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_BEDAT-HIGH").text = date - 90
session.findById("wnd[0]/usr/txtS_EAN11-LOW").text = ""
session.findById("wnd[0]/usr/txtS_EAN11-HIGH").text = ""
session.findById("wnd[0]/usr/txtS_IDNLF-LOW").text = ""
session.findById("wnd[0]/usr/txtS_IDNLF-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_LTSNR-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_LTSNR-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_AKTNR-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_AKTNR-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtS_SAISO-LOW").text = ""
session.findById("wnd[0]/usr/ctxtS_SAISO-HIGH").text = ""
session.findById("wnd[0]/usr/txtS_SAISJ-LOW").text = ""
session.findById("wnd[0]/usr/txtS_SAISJ-HIGH").text = ""
session.findById("wnd[0]/usr/txtP_TXZ01").text = ""
session.findById("wnd[0]/usr/txtP_NAME1").text = ""
session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "7008"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "7013"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "7020"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "7039"
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press