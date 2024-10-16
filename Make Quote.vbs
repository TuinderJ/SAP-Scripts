dim PRODUCTION
PRODUCTION = true

dim repairOrderNumber, customerNumber, payerNumber, shipToNumber

if not PRODUCTION then
   repairOrderNumber = ""
end if

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
   Set session2   = connection.Children(1)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

askForUserInput()
if repairOrderNumber <> "1" and repairOrderNumber <> "2" then
   openRepairOrder()
end if
makeQuote()


sub askForUserInput()
   if repairOrderNumber = "" then
      repairOrderNumber = inputBox("What is the RO number you have?" & vbCr & "1) Phone Quote" & vbCr & "2) Employee", "RO Number")
      if repairOrderNumber = "" then
         wScript.quit
      end if
   end if
end sub

sub openRepairOrder()
   session2.findById("wnd[0]").maximize
   session2.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
   session2.findById("wnd[0]/tbar[0]/btn[0]").press
   session2.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = repairOrderNumber
   session2.findById("wnd[0]").sendVKey 0
   ' Get customer info
   session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2066/btnGS_ORDER_SCREENS-SCARCP_ICON").press
   for row = 0 to 6
      dim partnerFunction, partner
      partnerFunction = session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2066/subSUBSCREEN_2066:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2210/ssubPARTNER_SUBSCR_AREA:/DBM/SAPLCU05:0100/ssubPARTNER_SD:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0," & row & "]").text
      partner =         session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2066/subSUBSCREEN_2066:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:2210/ssubPARTNER_SUBSCR_AREA:/DBM/SAPLCU05:0100/ssubPARTNER_SD:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1," & row & "]").text
      select case partnerFunction
         case "Sold-to party"
            customerNumber = partner
         case "Payer"
            payerNumber = partner
         case "Ship-to party"
            shipToNumber = partner
      end select
   next
   session2.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2066/btnGS_ORDER_SCREENS-SCARCP_ICON").press
end sub

sub makeQuote()
   session.findById("wnd[0]").maximize
   session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER01"
   session.findById("wnd[0]/tbar[0]/btn[0]").press
   session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP01"
   if repairOrderNumber = "1" then
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "200000"
   elseif repairOrderNumber = "2" then
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "152558"
   else
      session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = customerNumber
   end if
   session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
   session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
   session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
   session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
   session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
   session.findById("wnd[0]/tbar[1]/btn[13]").press
   
   ' Partner selection (if it pops up)
   on error resume next
   session.findById("wnd[1]/usr/lbl[4,4]").setFocus
   if not err.number then
      err.clear
      if not PRODUCTION then
         msgBox "Payer" & vbCr & payerNumber & vbCr & vbCr & "Ship-to" & vbCr & shipToNumber,, "Customer"
      end if
      selectPartner()
      session.findById("wnd[1]/tbar[0]/btn[0]").press
   end if
   err.clear
   on error goto 0

   if repairOrderNumber = "1" then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2065/subSUBSCREEN_2065:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3201/txt/DBM/VBAK_COM-BSTNK").text = "Phone Quote"
   elseif repairOrderNumber = "2" then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2065/subSUBSCREEN_2065:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3201/txt/DBM/VBAK_COM-BSTNK").text = inputBox("Who is this for?", "Employee Name")
   else
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2065/subSUBSCREEN_2065:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3201/txt/DBM/VBAK_COM-BSTNK").text = repairOrderNumber
   end if
   session.findById("wnd[0]").sendVKey 0
   session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
   session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "W"
end sub

sub selectPartner()
   dim value, payerOrShipTo, options, column
   payerOrShipTo = "Payer"
   on error resume next
   ' Payer only check
   session.findById("wnd[1]/usr/lbl[12,5]").setFocus
   if err.number = 0 then
      options = "Payer"
   end if
   err.clear
   ' Both check
   session.findById("wnd[1]/usr/lbl[15,5]").setFocus
   if err.number = 0 then
      options = "Payer and Ship To"
   end if
   err.clear
   ' Ship to only check
   if session.findById("wnd[1]/usr/lbl[4,4]").text = "SH" then
      options = "Ship To"
      payerOrShipTo = "Ship To"
   end if
   select case options
      case "Payer"
         column = "12"
      case "Payer and Ship To"
         column = "15"
      case "Ship To"
         column = "10"
   end select
   if payerOrShipTo = "Payer" then
      session.findById("wnd[1]").sendVKey 2
      ' Press the find button
      session.findById("wnd[1]/tbar[0]/btn[71]").press

      ' Type what you want, find it, click on it
      session.findById("wnd[2]/usr/txtRSYSF-STRING").text = payerNumber
      session.findById("wnd[2]/tbar[0]/btn[0]").press
      session.findById("wnd[3]/usr/lbl[4,2]").setFocus
      session.findById("wnd[3]").sendVKey 2
      session.findById("wnd[1]").sendVKey 2

      if options = "Payer and Ship To" then
         payerOrShipTo = "Ship To"
      end if
   end if
   if payerOrShipTo = "Ship To" then
      ' Press the find button
      session.findById("wnd[1]/tbar[0]/btn[71]").press

      ' Type your ship to number
      session.findById("wnd[2]/usr/txtRSYSF-STRING").text = shipToNumber
      session.findById("wnd[2]/tbar[0]/btn[0]").press
      session.findById("wnd[3]/usr/lbl[7,2]").setFocus
      session.findById("wnd[3]/usr/lbl[7,2]").caretPosition = 2
      session.findById("wnd[3]").sendVKey 2

      ' Click on it since you were taken there. No check needed
      session.findById("wnd[1]").sendVKey 2

      session.findById("wnd[1]/tbar[0]/btn[0]").press
   end if
end sub

function removeFrontZeros(value)
   if left(value, 1) = "0" then
      value = removeFrontZeros(right(value, len(value) - 1))
   end if
   removeFrontZeros = value
end function