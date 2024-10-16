set fso = createObject("Scripting.FileSystemObject")
dim filesList()
redim filesList(0)
main()
set fso = nothing

sub main()
  dim index, message, answer
  index = 0

  message = getFilesAt("C:\Users\tuinderj\OneDrive - Rush Enterprises\Documents\SAP\SAP GUI\OneDrive Scripts\Quotes")

  do while true
    answer = inputBox(message, "Select a script to run...")
    if not isNumeric(answer) or answer = "" then
      wScript.quit
    end if
    
    set shell = wScript.createObject("WScript.Shell")
    shell.run "cscript """ & filesList(answer - 1) & """", 0, true
    set shell = nothing
  loop
end sub

function getFilesAt(dirname)
  set rootDirectory = fso.getFolder(dirname)
  ' Files
  set files = rootDirectory.files
  for each file in files
    redim preserve filesList(index)
    filesList(index) = file
    getFilesAt = getFilesAt & index + 1 & ": " & file.name & vbCr
    index = index + 1
  next
  
  set rootDirectory = nothing
  set files = nothing
end function