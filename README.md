"# SPSoap" 

***VBA Class Libraries for SharePoint SOAP Communication***

- - -
VBA library to communicate with SharePoint without the usage of the outdated Web Services/SOAP Toolkit. Library is based on MSXML Core Services 3.0

### WSDL Class Generator

Originally I wanted to use perl to dynamically generate the VBA code from a wsdl, but finally ended up using .NET Web Services interface and reflection to generate the parsing tree to build the cls code. This has the disadvantage to contract the service within c# before being able to build the VBA code.

### VBA Integration

Either include SPSoap.cls in your project or start by using the Example project. Also ensure a reference to MS XML, v3.0 exists. The following sample code shows how to access a SharePoint List and Display some results.

### Sample code

```VB
Sub LoadItems()
 
    ' Initialize SPSoap
    Dim ws As New SPSoap
    Call ws.init("https://mysharepointsite.com")

   ' Resultset
   Dim x As MSXML2.IXMLDOMNodeList
   
   Set x = ws.wsm_GetListItems("MySharePointList", "", Nothing, Nothing, "", Nothing, "")
   
   Dim root As IXMLDOMElement
   Set root = x.item(0)
   
   Dim elements As Variant
   Set elements = root.getElementsByTagName("rs:data")
   
   Dim strQuery: strQuery = ".//z:row"
           
   Set Items = root.SelectNodes(strQuery)
   Debug.Print "No of Items: " & Items.Length
   Dim item As Variant
   For Each item In Items
       Debug.Print "Title: " & item.getAttribute("ows_Title")
   Next
   
End Sub
