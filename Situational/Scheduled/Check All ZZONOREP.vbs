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

Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Add()
Set objOutlook = CreateObject("Outlook.Application")
Set objEmail = objOutlook.CreateItem(0)

Dim i, j, intCellRow, strArrayBranches, boolDuplicate, strNewExcelFilePath, strEmailBody, strSignature, intTotalOrders
Dim ordersOver1Month()
strArrayBranches = Array("7008", "7013", "7020", "7039")
intCellRow = 1
strNewExcelFilePath = strDesktop & "\Aged PO's " & Month(Date) & "-" & Day(Date) & ".xlsx"
strEmailBody = "Here are the PO's over 30 days old for the region." & vbCr & "Take some time to review these and get them cleared out if you can." & vbCr & vbCr
strSignature = vbCr & vbCr & "Joshua Tuinder" & vbCr & "Rush Truck Leasing" & vbCr & "Mountain Region Inventory Control Supervisor" & vbCr & "379 W 66th Way" & vbCr & "Denver, CO" & vbCr & "O: (720) 292-5808" & vbCr & "C: (720) 413-1681"

For Each branch in strArrayBranches
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZONOREP"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = branch
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    'PO, Vendor, Date, P/N, Description, Number of entries

    i = 0
    j = 0
    Do While true
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell i,"BEDAT"
        If session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"BEDAT") = "" Then exit do End If
        If Err.Number <> 0 Then exit do End If
        If CDate(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"BEDAT")) < Date - 30 Then
            For k = 0 to j - 1
                If session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"VBELN") = ordersOver1Month(0, k) Then
                    boolDuplicate = true
                    ordersOver1Month(5, k) = ordersOver1Month(5, k) + 1
                End If
            Next
            If Not boolDuplicate Then
                Redim Preserve ordersOver1Month(5, j)
                ordersOver1Month(0, j) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"VBELN")
                ordersOver1Month(1, j) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"NAME1")
                ordersOver1Month(2, j) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"BEDAT")
                ordersOver1Month(3, j) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"MATNR")
                ordersOver1Month(4, j) = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(i,"MAKTX")
                ordersOver1Month(5, j) = 1
                j = j + 1
            Else
                boolDuplicate = false
            End If
        End If
        i = i + 1
    Loop
    If ordersOver1Month(0, 0) = "" Then
        intTotalOrders = 0
    Else
        intTotalOrders = UBound(ordersOver1Month,2) + 1
    End If
    strEmailBody = strEmailBody & "Branch: " & branch & vbTab & "Total PO's: " & intTotalOrders & vbCr & vbCr
    objExcel.Cells(intCellRow, 1).Value = "Branch: " & branch
    objExcel.Cells(intCellRow, 2).Value = "Total PO's: " & intTotalOrders
    intCellRow = intCellRow + 1
    objExcel.Cells(intCellRow, 1).Value = "PO"
    objExcel.Cells(intCellRow, 2).Value = "Vendor"
    objExcel.Cells(intCellRow, 3).Value = "Date opened"
    objExcel.Cells(intCellRow, 4).Value = "Material/Part Number"
    objExcel.Cells(intCellRow, 5).Value = "Description"
    objExcel.Cells(intCellRow, 6).Value = "Open Lines on PO"
    intCellRow = intCellRow + 1
    For k = 0 To UBound(ordersOver1Month,2)
        objExcel.Cells(intCellRow, 1).Value = ordersOver1Month(0, k)
        objExcel.Cells(intCellRow, 2).Value = ordersOver1Month(1, k)
        objExcel.Cells(intCellRow, 3).Value = ordersOver1Month(2, k)
        objExcel.Cells(intCellRow, 4).Value = ordersOver1Month(3, k)
        objExcel.Cells(intCellRow, 5).Value = ordersOver1Month(4, k)
        objExcel.Cells(intCellRow, 6).Value = ordersOver1Month(5, k)
        intCellRow = intCellRow + 1
    Next
    intCellRow = intCellRow + 1
    Redim ordersOver1Month(5, 0)
    Err.Clear
Next

objExcel.Columns("A:F").AutoFit
objWorkbook.SaveAs(strNewExcelFilePath)
objWorkbook.Close
objExcel.workbooks.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing

With objEmail
    .To = _
        "elliottr@rushenterprises.com; " &_
        "pilottem@rushenterprises.com; " &_
        "arreya@rushenterprises.com; " &_
        "martinezj15@rushenterprises.com; " &_
        "rings@rushenterprises.com; " &_
        "swatsenbergr@rushenterprises.com; " &_
        "JacksonC1@RushEnterprises.com; " &_
        "Arndtb@RushEnterprises.com; " &_
        "DavilaV@RushEnterprises.com"
    .CC = _
        "woodr@rushenterprises.com"
    .BCC = _
        "tuinderj@rushenterprises.com"
    .Subject = "Aged PO's " & Month(Date) & "/" & Day(Date)
    .Body = strEmailBody & strSignature
    .htmlBody = .htmlBody & "<img src=""C:\Users\tuinderj\OneDrive - Rush Enterprises\Pictures\leasing-logo.png"">"
    .Attachments.Add strNewExcelFilePath
    .Send
End With

Dim fileSys
Set fileSys = CreateObject("Scripting.FileSystemObject")
filesys.DeleteFile strNewExcelFilePath
Set fileSys = Nothing

Set objOutlook = Nothing
Set objEmail = Nothing

session.findById("wnd[0]").sendVKey 3
session.findById("wnd[0]").sendVKey 3