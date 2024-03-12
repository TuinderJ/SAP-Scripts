dim branches, userInput, userMessage, recipients, recipients7008, recipients7013, recipients7020, workbookName, htmlOutput
workbookName = "Monthly Cycle Count.XLSX"
recipients7008 = "JacksonC1@RushEnterprises.com; PilotteM@rushenterprises.com"
recipients7013 = "Arreya@rushenterprises.com; ramirezm4@RushEnterprises.com"
recipients7020 = "martinezj15@rushenterprises.com; Arreya@rushenterprises.com; Rings@RushEnterprises.com"

decideWhichBranchesToDo()
for each branch in branches
  pullInventoryAndFormat(branch)
  sendEmail(branch)
next


function decideWhichBranchesToDo()
  if msgBox("Would you like to do every branch?", vbYesNo, "All Branches") = vbYes then
    branches = Array("7008","7013","7020")
    exit function
  end if

  userInput = inputBox("Write which branch you would like to do first.", "First Branch")
  if userInput = "" then
    wScript.quit
  end if

  branches = Array(userInput)

  do while true
    userMessage = "If you want to add another branch, enter it here." & vbCr
    for each branch in branches
      userMessage = userMessage & branch & ", "
    next
    userInput = inputBox(userMessage, "More Branches")
    if userInput <> "" then
      redim preserve branches(uBound(branches) + 1)
      branches(uBound(branches)) = userInput
    else
      exit do
    end if
  loop

  userMessage = "Are all these choices correct?" & vbCr
  for each branch in branches
    userMessage = userMessage & branch & ", "
  next
  if msgBox(userMessage, vbYesNo, "Verify") = vbNo then
    wScript.quit
  end if
end function


function pullInventoryAndFormat(branch)
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

  ' Pull Bin Location
  session.findById("wnd[0]/tbar[0]/okcd").text = "/NZZBIN"
  session.findById("wnd[0]/tbar[0]/btn[0]").press
  session.findById("wnd[0]/usr/chkP_BIN").selected = true
  session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = ""
  session.findById("wnd[0]/usr/txtS_EMNFR-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = branch
  session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = "0001"
  session.findById("wnd[0]/usr/txtS_LGPBE-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MTART-LOW").text = ""
  session.findById("wnd[0]/usr/ctxtS_MATKL-LOW").text = ""
  session.findById("wnd[0]/tbar[1]/btn[8]").press
  session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select

  ' Save the file
  session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\Users\tuinderj\OneDrive - Rush Enterprises\Desktop\"
  session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = branch & " " & workbookName
  session.findById("wnd[1]/tbar[0]/btn[11]").press

  session.findById("wnd[0]").sendVKey 3
  session.findById("wnd[0]").sendVKey 3

  ' Run the Excel Macro
  WScript.sleep 5000
  set excel = getObject(,"Excel.Application")
  set workbook = excel.workbooks(branch & " " & workbookName)
  workbook.Application.Run "PERSONAL.XLSB!AutomaticCycleCount"

  msgBox "Please review before sending", vbOk, "Review"

  ' Save and quit
  Set WshShell = WScript.CreateObject("WScript.Shell")
  desktop = WshShell.SpecialFolders("Desktop")
  excelFilePath = desktop & "\" & branch & " " & workbookName
  ' workbook.SaveAs(excelFilePath)
  workbook.Close
  ' excel.workbooks.Close
  ' excel.Quit

  Set WshShell = nothing
  set workbook = nothing
  set excel = nothing
end function


function sendEmail(branch)
  select case branch
    case "7008"
      recipients = recipients7008
    case "7013"
      recipients = recipients7013
    case "7020"
      recipients = recipients7020
  end select

  Set outlook = CreateObject("Outlook.Application")
  Set email = outlook.CreateItem(0)

  Set WshShell = WScript.CreateObject("WScript.Shell")
  desktop = WshShell.SpecialFolders("Desktop")
  excelFilePath = desktop & "\" & branch & " " & workbookName

  ' Signature
  htmlOutput = _
    "<p>Here is this month's cycle count. Please fill it out and send it back to me.</p>" & vbCr & _
    "<p>If it wants to split columns on 2 pages, click ""Enable Editing"". That should fix it.</>" & vbCr & _
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
    "<br />" & vbCr & _
    "<img src=""C:\Users\tuinderj\OneDrive - Rush Enterprises\Pictures\leasing-logo.png"" />" & vbCr & _
    "</p>" & vbCr & _
    "</body>" & vbCr & _
    "</html>"

  With email
    .To = recipients
    '.CC = ""
    '.BCC = ""
    .Subject = "Monthly Cycle Count " & date
    .htmlBody = htmlOutput
    .Attachments.Add excelFilePath
    .Send
  End With

  Set WshShell = nothing
  Set outlook = nothing
  Set email = nothing
end function