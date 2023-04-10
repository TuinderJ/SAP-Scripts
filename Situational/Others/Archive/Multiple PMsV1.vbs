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

Dim truck()
Dim temp, i, x, pm, pm2, of, of2, dot, dot2, dryr, dryr2, verifyTrucks, trucksNotFound, trucksFound

i = 0
Do While True
    temp = InputBox("What is the unit number of the next truck?" & vbCr & "Don't use a -" & vbCr & "If you're done, leave it blank.")
    If temp = "" Then
        Exit Do
    End If
    Redim Preserve truck(i)
    truck(i) = temp
    i = i + 1
Loop

Do While True
    i = 0
    verifyTrucks = ""
    For Each unit in truck
        verifyTrucks = verifyTrucks & vbCr & i + 1 & ": " & unit
        i = i + 1
    Next

    If MsgBox("Are all of these entered properly?" & verifyTrucks, vbYesNo) =  vbNo Then
        i = InputBox("What number do you need to change?" & "If you need to cancel the whole thing, leave blank." & verifyTrucks)
        If i = "" Then
            WScript.Quit
        End If
        i = i - 1
        temp = InputBox("What is the new number for " & truck(i) & "?" & vbCr & "Leave blank to remove it.")
        If temp <> "" Then
            truck(i) = temp
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

For Each unit in truck
    session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/VSEARCH"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/subSUBSCREEN1:/DBM/SAPLVM05:2000/subSUBSCREEN1:/DBM/SAPLVM05:2200/ctxtZZUN-LOW").text = unit
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ssubSUBSCREEN1:/DBM/SAPLVM05:1100/tabsTABSTRIP/tabpSEARCHVM/ssubSUBSCREEN1:/DBM/SAPLVM05:1200/btnBUTTON").press
    
    Do
        If session.findById("wnd[0]/sbar").text = "No vehicles could be selected" Then
            trucksNotFound = trucksNotFound & vbCr & unit
            trucksFound = trucksFound & vbCr
            Exit Do
        Else
            trucksFound = trucksFound & vbCr & unit
            session.findById("wnd[0]/usr/tabsMAIN/tabpORDER").select
            session.findById("wnd[0]/usr/tabsMAIN/tabpORDER/ssubDETAIL_SUBSCR:/DBM/SAPLVM08:2001/cntlDOCKING_CONTROL_PROXY/shellcont/shell").clickLink "ZS20","Column01"
            session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
            session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1001"
            session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "12"
            session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "7039"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            Err.Clear

            'Header
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREV_MILEAGE").text
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZENGINEHOURS").text = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-ZZPREVENGHOURS").text
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/subSUB_ACCOUNTING:/DBM/SAPLORDER_UI:2204/cmb/DBM/VBAK_COM-ZZORD_STATUS").key = "ON-SITE/MOBILE REPAIR"
            session.findById("wnd[0]").sendVKey 0

            'Read intervals
            Err.Clear
            
            i = 0
            Do While True
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
                If Err.Number <> 0 Then
                    Exit Do
                End If
                If inStr(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"PM") Then
                    pm = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                    pm2 = pm & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                End If
                If inStr(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"OF") Then
                    If CDate(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")) <= Date Then
                        of = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                        of2 = of & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                    End If
                    If int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"SCOUNT")) <= int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2067/subSUBSCREEN_2067:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3200/txt/DBM/VBAK_COM-MILEAGE").text) Then
                        of = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                        of2 = of & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                    End If
                End If
                Err.Clear
                i = i + 1
            Loop

            Err.Clear
            i = 0
            Do While True
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").setCurrentCell i,"STYPE"
                If Err.Number <> 0 Then
                    Exit Do
                End If
                If inStr(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DOT") Then
                    If inStr(pm,"TPM") Then
                        If CDate(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")) <= (Date + 90) Then
                            dot = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                            dot2 = dot & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                        End If
                    Else
                        If CDate(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")) <= (Date + int(Right(pm,Len(pm) - 2))) Then
                            dot = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                            dot2 = dot & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                        End If
                    End If
                End If
                If inStr(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE"),"DRYR") Then
                    If inStr(pm,"TPM") Then
                        If CDate(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")) <= (Date + 90) Then
                            dryr = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                            dryr2 = dryr & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                        End If
                    Else
                        If CDate(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"DATNEXT")) <= (Date + int(Right(pm,Len(pm) - 2))) Then
                            dryr = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE")
                            dryr2 = dryr & " -" & session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB01/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA2:/DBM/SAPLORDER_UI:2063/subSUBSCREEN_2063:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLVSALES_UI:2000/subCONTAINER_LEGAL_INTERVALS:/DBM/SAPLVM07:1400/cntlIOBJ_MULTI_SINT/shellcont/shell").getCellValue(i,"STYPE_DESC")
                        End If
                    End If
                End If
                i = i + 1
            Loop

            'Make jobs
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02").select
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = pm2
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = pm
            session.findById("wnd[0]").sendVKey 0
            If of <> "" Then
                session.findById("wnd[0]").sendVKey 5
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = of2
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = of
                session.findById("wnd[0]").sendVKey 0
                of = 2
            End If
            If dot <> "" Then
                session.findById("wnd[0]").sendVKey 5
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = dot2
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = dot
                session.findById("wnd[0]").sendVKey 0
                If of = 2 Then
                    dot = 3
                Else
                    dot = 2
                End If
            End If
            If dryr <> "" Then
                session.findById("wnd[0]").sendVKey 5
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR1").text = dryr2
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB02/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI1:2341/txt/DBM/JOB_COM-DESCR2").text = dryr
                session.findById("wnd[0]").sendVKey 0
                If dot = 3 Then
                    dryr = 4
                ElseIf dot = 2 Then
                    dryr = 3
                Else
                    dryr = 2
                End If
            End If

            'Add Labor
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03").select
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM HD"
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = "1"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]").sendVKey 0

            If of <> "" Then
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "PM OF"
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = of
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 0
                of = ""
            End If

            If dot <> "" Then
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "0000010"
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = dot
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 0
                dot = ""
            End If

            If dryr <> "" Then
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-ITOBJID").text = "1310056"
                session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB03/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3310/ctxt/DBM/S_POS-JOBS").text = dryr
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 0
                dot = ""
            End If

            ' session.findById("wnd[0]").sendVKey 11
        End If
    Loop While false
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