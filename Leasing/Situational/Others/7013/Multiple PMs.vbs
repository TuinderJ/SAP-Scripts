Dim PRODUCTION, branch, advisorNumber
PRODUCTION = True
branch = "7013"
advisorNumber = "19126"

Function readIntervals()
  Err.Clear
  On Error Resume Next
  Redim serviceInterval(2, 0)
  i = 0
  x = 0

  Do While True
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    If Err.Number <> 0 Then
      Err.Clear
      Exit Do
    End If
    ' If PM
    If inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"PM") = 1 Then
      storeInterval()
    End If
    i = i + 1
  Loop
  i = 0
  Do While True
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
    If Err.Number <> 0 Then
        Exit Do
    End If
    ' If OF
    If inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"OF") = 1 Then
      If checkIfDue() = True Then
        storeInterval()
      Else
        workOrder.Add "Oil", session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
      End If
    End If
    ' If DOT, DRYR, RFPM, DEFFI
    If _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DOT") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DRYR") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"RFPM") = 1 Or _
    inStr(1, session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DEFFI") = 1 _
    Then
      If checkIfDue() = True Then
        ' MsgBox("It decided the interval is due")
        storeInterval()
      End If
    End If
    i = i + 1
  Loop
End Function

Function storeInterval()
  Redim Preserve serviceInterval(2, x)
  serviceInterval(0, x) = x + 1
  serviceInterval(1, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
  serviceInterval(2, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
  x = x + 1
End Function

Function checkIfDue()
  checkIfDue = False
  Dim intervalDueDate, intervalMileage, truckCurrentMileage
  intervalDueDate = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")
  intervalMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")
  truckCurrentMileage = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text
  ' MsgBox("intervalDueDate: " & intervalDueDate & vbCr & "currentDate: " & Date & vbCr & "intervalMileage: " & intervalMileage & vbCr & "truckCurrentMileage: " & truckCurrentMileage)
  ' MsgBox("Date check: " & Date + CInt(Right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) & vbCr & "Interval: " & CDate(intervalDueDate))
  If intervalDueDate <> "" Then
    If Date + CInt(Right(serviceInterval(1, 0),Len(serviceInterval(1, 0)) - 2)) > CDate(intervalDueDate) Then
      checkIfDue = True
    End If
  End If
  If checkIfDue = False Then
    If _
    "" = intervalMileage Or _
    0 = CLng(intervalMileage) _
    Then
    Else
      If CLng(intervalMileage) <= CLng(truckCurrentMileage) Then
        checkIfDue = True
      End If
    End If
  End If
End Function

Function makeJobs()
  ' Go to job tab
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
  For job = 0 to UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = serviceInterval(1, job) & " -" & serviceInterval(2, job)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = serviceInterval(1, job)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 5
  Next
End Function

Function addLabor()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
  For job = 0 To UBound(serviceInterval, 2)
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = serviceInterval(0, job)
    If inStr(1, serviceInterval(1, job),"PM") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM HD"
    End If
    If inStr(1, serviceInterval(1, job),"OF") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM OF"
    End If
    If inStr(1, serviceInterval(1, job),"DOT") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "10"
    End If
    If inStr(1, serviceInterval(1, job),"DRYR") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "1310056"
    End If
    If inStr(1, serviceInterval(1, job),"RFPM") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "9800050"
    End If
    If inStr(1, serviceInterval(1, job),"DEFFI") = 1 Then
      session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "4307002"
    End If
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
  Next
End Function

Function addParts()
  session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
End Function

function vehicleIsLocked()
  on error resume next
  ' Select the order tab and click on internal order
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
  session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = advisorNumber
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
  session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = branch
  session.findById("wnd[0]").sendVKey 0
  message = session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSBAL_DISPLAY:0101/cntlSAPLSBAL_DISPLAY_CONTAINER/shellcont/shell").getCellValue(0, "T_MSG")
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  session.findById("wnd[1]/tbar[0]/btn[0]").press
  if message <> "" then
    if msgBox(message & vbCr & "Would you like to try again?", vbYesNo, "Try Again") = 6 then
      vehicleIsLocked = true
      exit function
    else
      Wscript.Quit
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

Dim trucks()
Dim serviceInterval()
Dim userInput, i, x, verifyTrucks, trucksNotFound, trucksFound, result

' Ask the user which trucks they need
i = 0
Do While True
    userInput = replace(inputBox("What is the unit number of the next truck?" & vbCr & "If you're done, leave it blank."),"-", "")
    If userInput = "" Then
        Exit Do
    End If
    Redim Preserve trucks(i)
    trucks(i) = userInput
    i = i + 1
Loop

if trucks(0) = "" then
  WScript.Quit
end if

' Ask the user for verification
Do While True
    i = 0
    verifyTrucks = ""
    For Each truck in trucks
        verifyTrucks = verifyTrucks & vbCr & i + 1 & ": " & truck
        i = i + 1
    Next

    If MsgBox("Are all of these entered properly?" & verifyTrucks, vbYesNo) =  vbNo Then
        i = InputBox("What number do you need to change?" & "If you need to cancel the whole thing, leave blank." & verifyTrucks)
        If i = "" Then
            WScript.Quit
        End If
        i = i - 1
        userInput = InputBox("What is the new number for " & truck(i) & "?" & vbCr & "Leave blank to remove it.")
        If userInput <> "" Then
            truck(i) = userInput
        Else
            If i = UBound(truck) Then
                Redim Preserve truck(i - 1)
            End If
            If i <= UBound(truck)Then
                Do Until i => UBound(truck)
                    truck(i) = truck(i + 1)
                    i = i + 1
                Loop
                Redim Preserve truck(i - 1)
            End If
        End If
    Else
        Exit Do
    End If
Loop

For Each truck in trucks
  On Error Resume Next
  session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = truck
  session.findById("wnd[0]").sendVKey 0
  session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
  
  Do
    If session.findById("wnd[0]/sbar").text = "No vehicles could be selected" Then
      trucksNotFound = trucksNotFound & vbCr & truck
      trucksFound = trucksFound & vbCr
      Exit Do
    end if
    trucksFound = trucksFound & vbCr & truck
    do while vehicleIsLocked()
    loop
    Err.Clear

    ' Header
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press

    readIntervals()
    makeJobs()
    addLabor()
    addParts()
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[1]/btn[37]").press
    session.findById("wnd[0]/tbar[1]/btn[13]").press
    session.findById("wnd[1]/usr/btnBUTTON_1").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
  Loop While false
  On Error Goto 0
Next

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsCNT_TAB/tabpTAB_01/ssubSEARCH_SUBSCREEN:/DBM/SAPLORDER_UI:1001/btnBUTTON04").press

If trucksNotFound <> "" Then
    MsgBox("These trucks were not found in SAP." & trucksNotFound)
End If
If trucksFound <> "" Then
    MsgBox("These trucks were created successfully." & trucksFound)
End If

Set objExcel = Nothing
Set objWorkbook = Nothing