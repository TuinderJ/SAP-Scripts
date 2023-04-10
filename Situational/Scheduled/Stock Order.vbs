Dim currentLine, tempPartStorage, i, x, ro
Dim row()
Dim part()
Dim part2()

Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\Main Folder\Parts\Parts Main.csv",1)
i = 0
currentLine = objFileToRead.ReadLine()
do while not objFileToRead.AtEndOfStream
    currentLine = objFileToRead.ReadLine()
    tempPartStorage = split(currentLine,",")
    If tempPartStorage(1) <> "0" Then
        Redim Preserve part(1, i)
        part(0, i) = tempPartStorage(1)
        part(1, i) = tempPartStorage(2)
        i = i + 1
    End If
loop
objFileToRead.Close
Set objFileToRead = Nothing

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

session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER01"
session.findById("wnd[0]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP21"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "7039"
session.findById("wnd[0]/usr/chk/DBM/ORDER_CREATION-NO_VEHICLE").selected = True
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "D"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text = "Stock Order"

'Put parts in SO
i = 0
Do until i > UBound(part, 2)
    'P/N
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = part(1, i)
    'Qty
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = part(0, i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").sendVKey 0
    i = i + 1
loop

'Delete 0's
Err.Clear
i = 0
x = 0
Do until Err.Number <> 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i, "KBMENG"
    if int(session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(i,"KBMENG")) = 0 Then
        if Err.Number <> 0 Then
            Err.Clear
            session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i + 1, "KBMENG"
        Else
            Redim Preserve row(x)
            Redim Preserve part2(1, x)
            row(x) = i
            part2(0, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(i,"ZMENG")
            part2(1, x) = session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").getCellValue(i,"ITOBJID")
            x = x + 1
        End if
    End If
    i = i + 1
Loop

Err.Clear
i = CStr(row(0))
x = 1
Do until Err.Number <> 0
    i = i + ", " + CStr(row(x))
    x = x + 1
Loop
Err.Clear

session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").selectedRows = i
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressToolbarButton "ITEM_DELETE"
session.findById("wnd[0]").sendVKey 11

'Set jobs to 0
i = 0
Do until Err.Number <> 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i, "JOBS"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell i,"JOBS","0"
    if Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i + 1, "KBMENG"
    End if
    i = i + 1
Loop

session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressEnter

session.findById("wnd[0]").sendVKey 11
ro = Right(Left(session.findById("wnd[0]/titl").text,40),8)
Err.Clear

'New SO for back order
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER01"
session.findById("wnd[0]/tbar[0]/btn[0]").press

session.findById("wnd[0]/usr/cmb/DBM/ORDER_CREATION-AUFART").key = "ZP21"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PARTNER").text = "7039"
session.findById("wnd[0]/usr/chk/DBM/ORDER_CREATION-NO_VEHICLE").selected = True
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-PERNR").text = "73363"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VKORG").text = "1000"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-VTWEG").text = "13"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-SPART").text = "99"
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_CREATION-WERKS").text = "1401"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/cmb/DBM/VBAK_COM-ZZWILL_CALL_DEL").key = "D"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/VBAK_COM-BSTNK").text = "Stock Back Order"

'Put parts in new SO
i = 0
Do until i > UBound(part2,2)
    'P/N
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/ctxt/DBM/S_POS-ITOBJID").text = part2(1, i)
    'Qty
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA1:/DBM/SAPLORDER_UI:2061/subSUBSCREEN_2061:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3321/txt/DBM/S_POS-ZMENG").text = part2(0, i)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]").sendVKey 0
    i = i + 1
loop

session.findById("wnd[0]").sendVKey 11

'Return to old SO and put X in Xfer box
session.findById("wnd[0]/tbar[0]/okcd").text = "/N/DBM/ORDER03"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxt/DBM/ORDER_SEARCH-VBELN").text = ro
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04").select
session.findById("wnd[0]/tbar[1]/btn[13]").press
session.findById("wnd[1]").sendVKey 0

Err.Clear
i = 0
Do until Err.Number <> 0
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i, "JOBS"
    session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").modifyCell i,"ZZXFER_RTC","x"
    if Err.Number <> 0 Then
        Err.Clear
        session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").setCurrentCell i + 1, "KBMENG"
    End if
    i = i + 1
Loop
session.findById("wnd[0]/usr/ssubORDER_SUBSCREEN:/DBM/SAPLATAB:0100/tabsTABSTRIP100/tabpTAB04/ssubSUBSC:/DBM/SAPLATAB:0200/subAREA3:/DBM/SAPLORDER_UI:2071/subSUBSCREEN_2071:/DBM/SAPLORDER_UI:2048/subSUBSCREEN:/DBM/SAPLORDER_UI:3400/cntl2400_CUSTOM_CONTAINER3400/shellcont/shell").pressEnter