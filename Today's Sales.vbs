dim fileSystemObject
set fileSystemObject = createObject("Scripting.FileSystemObject")
include("Z:\utilities.vbs")

main()
set fileSystemObject = nothing


sub include (file)
	'Create objects for opening text file
	set fso = createObject("Scripting.FileSystemObject")
	set textFile = fso.openTextFile(file, 1)

	'Execute content of file.
	executeGlobal textFile.readAll

	'CLose file
	textFile.close

	'Clean up
	set fso = nothing
	set textFile = nothing
end sub

sub main()
  goToTCode("zsalesrep")
  session.findById("wnd[0]/usr/ctxtP_ERDAT1").text = "10/12/2024"
  session.findById("wnd[0]/usr/ctxtS_WERKS1-LOW").text = "1401"
  session.findById("wnd[0]/tbar[1]/btn[8]").press
end sub