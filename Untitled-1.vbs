Dim customerPricingData(), vehiclePricingData(), customerNumberChangeData()

Function loadConfigFile()
  'Find the path name of this script
  strPath = Wscript.scriptFullname
  'Create File System Object
  Set objFSO = createObject("Scripting.FileSystemObject")
  'Create object for this script's file
  Set objFile = objFSO.getFile(strPath)
  'Get the folder that this script is located in
  strFolder = objFSO.getParentFolderName(objFile)
  'Create an Excel Object
  Set objExcel = createObject("Excel.Application")
  'Open the Rebill Pricing Excel File
  Set objWorkbook = objExcel.workBooks.open(strFolder & "\Rebill Pricing.xlsx")
  '-----------------------------Customer Pricing-----------------------------
  'Load the sheet and store the data
  Set objCustomerPricingSheet = objWorkbook.worksheets("Customer Pricing")
  Dim rowcount
  rowcount = objCustomerPricingSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve customerPricingData(4, i - 2)
    customerPricingData(0, i - 2) = objCustomerPricingSheet.cells(i, 2)
    customerPricingData(1, i - 2) = objCustomerPricingSheet.cells(i, 3)
    customerPricingData(2, i - 2) = objCustomerPricingSheet.cells(i, 4)
    customerPricingData(3, i - 2) = objCustomerPricingSheet.cells(i, 5)
  Next

  'End of data, clear memory
  Set objCustomerPricingSheet = Nothing

  '-----------------------------Vehicle Pricing-----------------------------
  'Load the sheet and store the data
  Set objVehiclePricingSheet = objWorkbook.worksheets("Vehicle Pricing")
  Dim rowcount
  rowcount = objVehiclePricingSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve vehiclePricingData(4, i - 2)
    vehiclePricingData(0, i - 2) = objVehiclePricingSheet.cells(i, 1)
    vehiclePricingData(1, i - 2) = objVehiclePricingSheet.cells(i, 2)
    vehiclePricingData(2, i - 2) = objVehiclePricingSheet.cells(i, 3)
    vehiclePricingData(3, i - 2) = objVehiclePricingSheet.cells(i, 4)
  Next

  'End of data, clear memory
  Set objVehiclePricingSheet = Nothing

  '-----------------------------Customer Number Change-----------------------------
  'Load the sheet and store the data
  Set objCustomerNumberChangeSheet = objWorkbook.worksheets("Customer Number Change")
  Dim rowcount
  rowcount = objCustomerNumberChangeSheet.Usedrange.Rows.Count

  For i = 2 To rowcount
    Redim Preserve customerNumberChangeData(1, i - 2)
    customerNumberChangeData(0, i - 2) = objCustomerNumberChangeSheet.cells(i, 2)
    customerNumberChangeData(1, i - 2) = objCustomerNumberChangeSheet.cells(i, 3)
  Next

  'End of data, clear memory
  Set objCustomerNumberChangeSheet = Nothing

  objWorkbook.close
  objExcel.workbooks.close
  objExcel.quit

  Set objWorkbook = Nothing
  Set objExcel = Nothing
  Set objFile = Nothing
  Set objFSO = Nothing
End Function