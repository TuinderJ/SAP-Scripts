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
If Not IsObject(session2) Then
  Set session2    = connection.Children(1)
End If
If IsObject(WScript) Then
  WScript.ConnectObject session,     "on"
  WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

' Go to Reservation List
session.findById("wnd[0]/tbar[0]/okcd").text = "/NMB24"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "7039"
session.findById("wnd[0]/tbar[1]/btn[8]").press

dim repairOrder, order
i = 0
do while true
  if order <> getOrderNumber(i) then
    getAndGoToRepairOrder()
    if msgbox("Do you want to continue?", vbYesNo, "Continue") = vbNo then
      exit do
    end if
  end if
  order = getOrderNumber(i)
  i = i + 1
loop

function getAndGoToRepairOrder()
  goToRow(i)
  pressItemButton()
  repairOrder = getRepairOrderNumber()
  openRepairOrderOnScreen2(repairOrder)
  goBackOnePage()
end function

function getOrderNumber(desiredRow)
  getOrderNumber = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").getCellValue(desiredRow,"AUFNR_H")
end function

function pressItemButton()
  session.findById("wnd[0]/tbar[1]/btn[17]").press
end function

function getRepairOrderNumber()
  getRepairOrderNumber = split(session.findById("wnd[0]/usr/txtRESB-SGTXT").text," ")(1)
end function

function goBackOnePage()
  session.findById("wnd[0]/tbar[0]/btn[3]").press
end function

function goToRow(desiredRow)
  session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell desiredRow,"BDTER_I"
end function

function openRepairOrderOnScreen2(desiredRepairOrder)
  session2.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
  session2.findById("wnd[0]/tbar[0]/btn[0]").press
  session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = desiredRepairOrder
  session2.findById("wnd[0]").sendVKey 0
  session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
  session2.findById("wnd[0]/tbar[1]/btn[13]").press
end function