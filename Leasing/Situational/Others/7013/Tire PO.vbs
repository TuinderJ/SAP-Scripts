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

dim vendorNumber, unitNumber, roType, notes, headerText, branch, jobDescription
branch = "7013"

function askForUserInput()
 askForUserInput = false
 
 if vendorNumber = "" then
   vendorNumber = inputBox("What vendor is this for?" & vbCr & "1) Michelin" & vbCr & "2) Purcell Tire" & vbCr & "3) Bill Williams/Broadway Motors" & vbCr & "4) Paccar Parts Fleet Services" & vbCr & "5) Paclease", "Vendor ")
   if vendorNumber = "" then
     WScript.Quit
   elseif vendorNumber = "1" then
       vendorNumber = "214567"
   elseif vendorNumber = "2" then
       vendorNumber = "205145"
   elseif vendorNumber = "3" then
       vendorNumber = "215188"
   elseif vendorNumber = "4" then
       vendorNumber = "205174"
   elseif vendorNumber = "5" then
       vendorNumber = "200521"
   else
     vendorNumber - ""
     msgBox "Please enter a valid option.", 0, "Error"
     exit function
   end if
 end if

 if unitNumber = "" then
   unitNumber = replace(inputBox("What is the unit number for this PO?", "Unit Number"),".","")
   if unitNumber = "" then
     WScript.Quit
   end if
 end if

 if roType = "" then
   roType = inputBox("What type of RO will this be?" & vbCr & "1) Internal" & vbCr & "2) Retail" & vbCr & "3) VIO" & vbCr & "4) Other (Please leave notes)", "RO Type")
   if roType = "" then
     WScript.Quit
   elseif roType = "1" then
     roType = "Internal"
   elseif roType = "2" then
     roType = "Retail"
   elseif roType = "3" then
     roType = "VIO"
   elseif roType = "4" then
     roType = "Other"
   else
     roType = ""
     msgBox "Please enter a valid option.", 0, "Error"
     exit function
   end if
 end if

if jobDescription = "" then
  jobDescription = inputBox("What would you like the job name to be?", "Job Name")
  if jobDescription = "" then
    WScript.Quit
  end if
end if


 if notes = "" then
   notes = inputBox("What notes would you like to leave?" & vbCr & "If nothing, just leave this blank.", "Notes")
 end if
 
 ' If all input is received, return true to move on
 askForUserInput = true
end function

function findItteration()
  on error resume next
  dim test, a, b
  a = 0
  b = 0
  do while true
     test = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & b & a & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,0]").text
     if err.number = 0 then
       exit do
     end if
     err.clear
     a = a + 1
     if a = 10 then
        a = 0
        b = b + 1
     end if
  loop
  findItteration = b & a
end function

' Header text breakdown
' Vendor|Unit_Number|RO_Type & vbCr & _
' Notes
do until askForUserInput()
loop
headerText = vendorNumber & "|" & unitNumber & "|" & roType & vbCr & jobDescription & vbCr & "Don't change anything above this line" & vbCr & "Notes: " & notes

'Go to ME21N
session.findById("wnd[0]/tbar[0]/okcd").text = "/NME21N"
session.findById("wnd[0]/tbar[0]/btn[0]").press

'Header
itteration = findItteration()
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/cmbMEPO_TOPLINE-BSART").key = "ZSER"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").text = vendorNumber
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3").select
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT3/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = headerText

'Line item
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]").text = "K"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[4,0]").text = "Tires - " & unitNumber
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[5,0]").text = "1"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[6,0]").text = "EA"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[7,0]").text = "1"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:00" & itteration & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[9,0]").text = "1052"
session.findById("wnd[0]").sendVKey 0

'G/L Account
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/ctxtMEACCT1100-SAKTO").text = "613000"
session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0020/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = branch & "00"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]/tbar[0]/okcd").text = "/NME23N"
session.findById("wnd[0]/tbar[0]/btn[0]").press
