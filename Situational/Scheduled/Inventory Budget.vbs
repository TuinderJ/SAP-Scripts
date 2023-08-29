function askForRecipients()
    answer = inputBox("Who do you want to send this to?" & vbCr & "1) You" & vbCr & "2) All Branches")
    if answer = "1" Or answer = "" Then
        recipients = "TuinderJ@rushenterprises.com"
    elseif answer = "2" Then
        recipients = _
        "elliottr@rushenterprises.com; " &_
        "pilottem@rushenterprises.com; " &_
        "arreya@rushenterprises.com; " &_
        "martinezj15@rushenterprises.com; " &_
        "rings@rushenterprises.com; " &_
        "swatsenbergr@rushenterprises.com; " &_
        "JacksonC1@RushEnterprises.com; " &_
        "TuinderJ@rushenterprises.com;" &_
        "DavilaV@RushEnterprises.com;" &_
        "SheridanM@RushEnterprises.com"
    else
        WScript.Quit
    end if
end function

Function pullBinLocationReport()
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "HC-2710B-RPS16-P1:APE"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/btn%_S_WERKS_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "7008"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "7013"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "7020"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "7039"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/btn%_S_LGORT_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV").select
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "0003"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "EXCOST"
    session.findById("wnd[0]/tbar[1]/btn[30]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WERKS"
    session.findById("wnd[0]/tbar[1]/btn[42]").press

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "Inventory vs Budget.XLSX"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]").sendVKey 3
    session.findById("wnd[0]").sendVKey 3
End Function


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

Dim strFilePath, strArrayBranches, htmlOutput, answer, recipients
Dim branchTargets(1, 3)
Dim branchInventory(1, 3)

branchTargets(0, 0) = 7008
branchTargets(1, 0) = 85000
branchTargets(0, 1) = 7013
branchTargets(1, 1) = 35000
branchTargets(0, 2) = 7020
branchTargets(1, 2) = 15000
branchTargets(0, 3) = 7039
branchTargets(1, 3) = 45000

Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
strFilePath = strDesktop & "\Inventory vs Budget.XLSX"

Dim fileSys
Set fileSys = CreateObject("Scripting.FileSystemObject")
filesys.DeleteFile strFilePath
Set fileSys = Nothing
On Error Goto 0
Err.Clear

askForRecipients()
pullBinLocationReport()

WScript.Sleep 1500
Set objExcel = GetObject(,"Excel.Application")
WScript.Sleep 1500
Set objWorkbook = objExcel.Workbooks("Inventory vs Budget.XLSX")
objWorkbook.Application.Run "PERSONAL.XLSB!InventoryVsBudget"

For i = 0 To UBound(branchTargets, 2)
    For j = 0 To UBound(branchTargets, 1)
        branchInventory(j, i) = objExcel.Cells(j + 1, i + 12).Value
    Next
Next

htmlOutput = _
    "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 3.2//EN"">" & vbCr & _
    "<html>" & vbCr & _
    "<head>" & vbCr & _
    "<meta name=""Generator"" content=""MS Exchange Server version 16.0.13801.20804""/>" & vbCr & _
    "</head>" & vbCr & _
    "<body>" & vbCr & _
    "<p>" & vbCr & _
    "<font face=""Calibri"">" & vbCr & _
    "Here is how much budget is left for our inventories." & vbCr & _
    "<br />" & vbCr & _
    "<br />" & vbCr & _
    "Branch: " & branchTargets(0, 0) & vbCr & _
    "<br />" & vbCr & _
    "Target: $" & FormatNumber(branchTargets(1, 0), 2) & vbCr & _
    "<br />" & vbCr & _
    "Current: $" & FormatNumber(branchInventory(1, 0), 2) & vbCr & _
    "<br />" & vbCr & _
    "Remaining: $"
if FormatNumber(branchTargets(1, 0) - branchInventory(1, 0), 2) > 0 Then
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Green"">" & FormatNumber(branchTargets(1, 0) - branchInventory(1, 0), 2) & "</span>"
Else
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Red"">" & FormatNumber(branchTargets(1, 0) - branchInventory(1, 0), 2) & "</span>"
End If
htmlOutput = htmlOutput + _
    vbCr & _
    "<br />" & vbCr & _
    "<br />" & vbCr & _
    "Branch: " & branchTargets(0, 1) & vbCr & _
    "<br />" & vbCr & _
    "Target: $" & FormatNumber(branchTargets(1, 1), 2) & vbCr & _
    "<br />" & vbCr & _
    "Current: $" & FormatNumber(branchInventory(1, 1), 2) & vbCr & _
    "<br />" & vbCr & _
    "Remaining: $"
if FormatNumber(branchTargets(1, 1) - branchInventory(1, 1), 2) > 0 Then
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Green"">" & FormatNumber(branchTargets(1, 1) - branchInventory(1, 1), 2) & "</span>"
Else
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Red"">" & FormatNumber(branchTargets(1, 1) - branchInventory(1, 1), 2) & "</span>"
End If
htmlOutput = htmlOutput + _
    "<br />" & vbCr & _
    "<br />" & vbCr & _
    "Branch: " & branchTargets(0, 2) & vbCr & _
    "<br />" & vbCr & _
    "Target: $" & FormatNumber(branchTargets(1, 2), 2) & vbCr & _
    "<br />" & vbCr & _
    "Current: $" & FormatNumber(branchInventory(1, 2), 2) & vbCr & _
    "<br />" & vbCr & _
    "Remaining: $"
if FormatNumber(branchTargets(1, 2) - branchInventory(1, 2), 2) > 0 Then
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Green"">" & FormatNumber(branchTargets(1, 2) - branchInventory(1, 2), 2) & "</span>"
Else
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Red"">" & FormatNumber(branchTargets(1, 2) - branchInventory(1, 2), 2) & "</span>"
End If
htmlOutput = htmlOutput + _
    "<br />" & vbCr & _
    "<br />" & vbCr & _
    "Branch: " & branchTargets(0, 3) & vbCr & _
    "<br />" & vbCr & _
    "Target: $" & FormatNumber(branchTargets(1, 3), 2) & vbCr & _
    "<br />" & vbCr & _
    "Current: $" & FormatNumber(branchInventory(1, 3), 2) & vbCr & _
    "<br />" & vbCr & _
    "Remaining: $"
if FormatNumber(branchTargets(1, 3) - branchInventory(1, 3), 2) > 0 Then
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Green"">" & FormatNumber(branchTargets(1, 3) - branchInventory(1, 3), 2) & "</span>"
Else
    htmlOutput = htmlOutput + _
    "<span style = ""Color: Red"">" & FormatNumber(branchTargets(1, 3) - branchInventory(1, 3), 2) & "</span>"
End If
htmlOutput = htmlOutput + _
    "</font>" & vbCr & _
    "</p>" & vbCr & _
    "<br />" & vbCr & _
    "<br />" & vbCr & _
    "<p>" & vbCr & _
    "<font face=""Calibri"">" & vbCr & _
    "<b>Joshua Tuinder" & vbCr & _
    "<br />" & vbCr & _
    "Rush Truck Leasing" & vbCr & _
    "<br />" & vbCr & _
    "Mountain Region Inventory Control Supervisor</b>" & vbCr & _
    "<br />" & vbCr & _
    "379 W 66th Way" & vbCr & _
    "<br />" & vbCr & _
    "Denver, CO" & vbCr & _
    "<br />" & vbCr & _
    "O: (720) 292-5808" & vbCr & _
    "<br />" & vbCr & _
    "C: (720) 413-1681" & vbCr & _
    "</font>" & vbCr & _
    "<br />" & vbCr & _
    "<img src=""C:\Users\tuinderj\OneDrive - Rush Enterprises\Pictures\leasing-logo.png"" />" & vbCr & _
    "</p>" & vbCr & _
    "</body>" & vbCr & _
    "</html>"

objWorkbook.Save
objWorkbook.Close
Set objExcel = Nothing
Set objWorkbook = Nothing

Set objOutlook = CreateObject("Outlook.Application")
Set objEmail = objOutlook.CreateItem(0)

With objEmail
    .To = recipients
    '.CC = ""
    '.BCC = ""
    .Subject = "Inventory Targets " & Month(Date) & "/" & Day(Date)
    .htmlBody = htmlOutput
    .Send
End With

Set objOutlook = Nothing
Set objEmail = Nothing