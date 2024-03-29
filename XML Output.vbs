Dim currentNode
 
Set xmlParser = CreateObject("Msxml2.DOMDocument")
 
'Creating an XML declaration
xmlParser.appendChild(xmlParser.createProcessingInstruction("xml","version = '1.0' encoding = 'windows-1251'"))
 
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

'Maximize the SAP window
session.findById("wnd[0]").maximize

enumeration "wnd[0]"

' MsgBox "Done!",VbSystemModal Or vbInformation

Sub enumeration(SAPRootElementId)
  Set SAPRootElement = session.findById(SAPRootElementId)

  'Creating a root element
  Set XMLRootNode = xmlParser.appendChild(xmlParser.createElement(SAPRootElement.Type))

  enumChildrens SAPRootElement, XMLRootNode
  
  xmlParser.save("C:\Users\tuinderj\OneDrive - Rush Enterprises\Documents\SAP_tree.xml")
  CreateObject("WScript.Shell").Run("""C:\Users\tuinderj\OneDrive - Rush Enterprises\Documents\SAP_tree.xml""")
End Sub

Sub enumChildrens(SAPRootElement, XMLRootNode)
  For i = 0 To SAPRootElement.Children.Count - 1
    Set SAPChildElement = SAPRootElement.Children.ElementAt(i)
    
    'Create a node
    Set XMLSubNode = XMLRootNode.appendChild(xmlParser.createElement(SAPChildElement.Type))
    
    'Attribute Name
    Set attrName = xmlParser.createAttribute("Name")
    attrName.Value = SAPChildElement.Name
    XMLSubNode.setAttributeNode(attrName)
    
    'Attribute Text
    If(Len(SAPChildElement.Text)> 0) Then
      Set attrText = xmlParser.createAttribute("Text")
      attrText.Value = SAPChildElement.Text
      XMLSubNode.setAttributeNode(attrText)
    End If
    
    'Attribute Id
    Set attrId = xmlParser.createAttribute("Id")
    attrId.Value = SAPChildElement.Id
    XMLSubNode.setAttributeNode(attrId)
    
    'If the current object is a container, then iterate through the child elements
    If(SAPChildElement.ContainerType) Then enumChildrens SAPChildElement, XMLSubNode
  Next
End Sub