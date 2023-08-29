Dim PRODUCTION, BRANCH, ADVISORNUMBER
PRODUCTION = True
BRANCH = "7039"
ADVISORNUMBER = "73363"

function vehicleIsLocked()
  on error resume next
  ' Select the order tab and click on internal order
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = ADVISORNUMBER
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = BRANCH
  session.findById("wnd[0]").sendVKey 0
  message = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0, "T_MSG")
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  if message <> "" then
    if msgBox(message & vbCr & "Would you like to try again?", vbYesNo, "Try Again") = 6 then
      vehicleIsLocked = true
      exit function
    else
      userWishesToContinue = false
      vehicleIsLocked = false
      redim preserve trucksNotCreated(trucksNotCreatedCount)
      trucksNotCreated(trucksNotCreatedCount) = customerUnitNumber
      trucksNotCreatedCount = trucksNotCreatedCount + 1
      exit function
    end if
  end if
  vehicleIsLocked = false
end function

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

On Error Resume Next
session.findById("wnd[0]").maximize

Dim unitNumbers()
Dim defQuantities()
Dim dates()
Dim initials()
Dim userInput, i, x, verifyTrucks, trucksNotFound, trucksFound, result, listSoFar

i = 0
Do While True
  ' listSoFar = ""
  ' For counter = 0 To i
  '   listSoFar = listSoFar & vbCr & initials(counter) & "    " & dates(counter) & "    " & unitNumbers(counter) & "    " & defQuantities(counter)
  ' Next
  userInput = UCase(InputBox("What are the initials of who dispensed the DEF?" & vbCr & "If you're done, leave it blank." & vbCr & listSoFar, "Initials"))
  If userInput = "" Then
      Exit Do
  End If
  Redim Preserve initials(i)
  initials(i) = userInput
  listSoFar = listSoFar & vbCr & userInput
  Redim Preserve dates(i)
  userInput = InputBox("What is the date that the DEF was used?" & vbCr & listSoFar, "Date")
  dates(i) = userInput
  listSoFar = listSoFar & "      " & userInput
  Redim Preserve unitNumbers(i)
  userInput = InputBox("What is the unit number of the truck?" & vbCr & listSoFar, "Unit Number")
  unitNumbers(i) = userInput
  listSoFar = listSoFar & "      " & userInput
  Redim Preserve defQuantities(i)
  userInput = InputBox("How many gallons were used?" & vbCr & listSoFar, "Qty")
  defQuantities(i) = userInput
  listSoFar = listSoFar & "      " & userInput
  i = i + 1  
Loop
i = i - 1

If MsgBox("Do you want to proceed?" & vbCr & listSoFar, vbYesNo, "Verify") = vbNo Then
  WScript.Quit
End If

If i = -1 Then
  WScript.Quit
End If

For counter = 0 To i
  On Error Resume Next
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unitNumbers(counter)
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
  
  Do
    If session.findById("wnd[0]/sbar").text = "No vehicles could be selected" Then
      trucksNotFound = trucksNotFound & vbCr & unitNumber(counter)
      trucksFound = trucksFound & vbCr
      Exit Do
    Else
      trucksFound = trucksFound & vbCr & unitNumber(counter)
      ' Make RO
      Do While vehicleIsLocked()
      Loop

      ' Header
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/btnCNT_BTN_HEADTEXT").press
      session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "DEF"
      session.findById("wnd[1]/tbar[0]/btn[8]").press

      ' Make Job
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = ("DEF " & dates(counter))
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/btnJOB_LONG_TEXT").press
      session.findById("wnd[1]/usr/cntlLTEXT_CONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[0]/shell").text = "Initials: " & initials(counter)
      session.findById("wnd[1]/tbar[0]/btn[8]").press

      ' Add Part
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = "KLF030T:MBL"
      session.findById("wnd[0]").sendVKey 0
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = defQuantities(counter)
      session.findById("wnd[0]").sendVKey 0
      
      ' Save Release Save
      session.findById("wnd[0]").sendVKey 11
      session.findById("wnd[0]/tbar[1]/btn[37]").press
      session.findById("wnd[0]/tbar[1]/btn[37]").press
      session.findById("wnd[0]/tbar[1]/btn[13]").press
    End If
  Loop While false
  On Error Goto 0
Next

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press