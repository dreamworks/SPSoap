VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SPSoap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
' Configuration Settings
Private url As String
Private Const LOG As Boolean = False
Private Const Version As String = "1.0"

' Internal Variables
Private oReq As MSXML2.XMLHTTP


'
'
' Constructor
Private Sub class_initialize()
   ' Create MSXML2 XML HTTP Object
   Set oReq = CreateObject("MSXML2.XMLHTTP")
   If oReq Is Nothing Then
      Call MsgBox("could not initialize msxml2.xmlhttp", vbCritical, "Error")
   End If
End Sub


Public Sub init(site As String)
   url = site
End Sub


'***************************************************************************************
' Soap Body
'***************************************************************************************
Private Function Soap_Body(method As String, attrs As Collection) As Variant

    Dim Header As String
    Header = ""
    
    Header = Header & "<?xml version='1.0' encoding='utf-8'?>"
    Header = Header & "<soap:Envelope xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:soap='http://schemas.xmlsoap.org/soap/envelope/'>"
    Header = Header & "  <soap:Body>"
    Header = Header & "    <" & method & " xmlns='http://schemas.microsoft.com/sharepoint/soap/' >"
    
    If Not attrs Is Nothing Then
        Dim attr As Variant
        For Each attr In attrs
           Header = Header & "        <" & attr(0) & ">" & attr(1) & "</" & attr(0) & ">"
        Next attr
    End If
    
    Header = Header & "     </" & method & ">"
    Header = Header & "  </soap:Body>"
    Header = Header & "</soap:Envelope>"
    
    Soap_Body = Header
End Function

'
'
'
Public Function PrettyPrint(xml As String) As String
    Dim rdr As New SAXXMLReader
    Dim wrt As New MXXMLWriter
    
    Set rdr.contentHandler = wrt
    wrt.indent = True
    rdr.Parse (xml)
    PrettyPrint = wrt.output
End Function



Private Sub AddAttributes(ByRef attrs As Collection, key As String, varValue As Variant)
    
    Dim value As String
    value = ""
    
    
    If (VarType(varValue) = 9) Then
       If (varValue Is Nothing) Then
           Exit Sub
       Else
           ' it is IXMLDOMnodeList otherwise
           Dim x As IXMLDOMNodeList
           Set x = varValue
           value = x.item(0).xml
       End If
    
    Else
       value = CStr(varValue)
    End If
    
    
    ' Only Add to Collection if value is Set
    If Not (value = "") Then
       attrs.Add Array(key, value), key
    End If

End Sub


'
' Execute SOAP request
' & fetch result
Private Function Request(service As String, method As String, body As String) As IXMLDOMNodeList

   Set Request = Nothing
   
   On Error GoTo ErrRoutine
     
     oReq.Open "POST", url & service, False
     oReq.setRequestHeader "Content-Type", "text/xml"
     oReq.setRequestHeader "SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/" & method
     oReq.send body
        
     Dim Result As MSXML2.DOMDocument
     Set Result = oReq.responseXML
     
     Set Request = Result.SelectNodes("//soap:Envelope/soap:Body/" & method & "Response/" & method & "Result")
   
   
EndRoutine:
    Exit Function
    
       
ErrRoutine:
    
    If LOG Then
       MsgBox Err.Number & " - " & Err.description, vbOKOnly Or vbCritical, "SPSOAP"
    End If
    Set Request = Nothing
    
    
    GoTo EndRoutine

End Function

'****************************************************************
' Automatically generated Functions start here
'****************************************************************
Public Function wsm_GetVersionCollection(strlistID As String, strlistItemID As String, strFieldName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetVersionCollection"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistID", strlistID)
    Call AddAttributes(attrs, "strlistItemID", strlistItemID)
    Call AddAttributes(attrs, "strFieldName", strFieldName)

    body = Soap_Body(method, attrs)
    Set wsm_GetVersionCollection = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_AddAttachment(listName As String, listItemID As String, fileName As String, attachment As Variant) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddAttachment"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "listItemID", listItemID)
    Call AddAttributes(attrs, "fileName", fileName)
    Call AddAttributes(attrs, "attachment", attachment)

    body = Soap_Body(method, attrs)
    Set wsm_AddAttachment = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetAttachmentCollection(listName As String, listItemID As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetAttachmentCollection"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "listItemID", listItemID)

    body = Soap_Body(method, attrs)
    Set wsm_GetAttachmentCollection = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_DeleteAttachment(listName As String, listItemID As String, url As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "DeleteAttachment"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "listItemID", listItemID)
    Call AddAttributes(attrs, "url", url)

    body = Soap_Body(method, attrs)
    Set wsm_DeleteAttachment = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_CheckOutFile(pageUrl As String, checkoutToLocal As String, lastmodified As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "CheckOutFile"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "pageUrl", pageUrl)
    Call AddAttributes(attrs, "checkoutToLocal", checkoutToLocal)
    Call AddAttributes(attrs, "lastmodified", lastmodified)

    body = Soap_Body(method, attrs)
    Set wsm_CheckOutFile = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UndoCheckOut(pageUrl As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "UndoCheckOut"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "pageUrl", pageUrl)

    body = Soap_Body(method, attrs)
    Set wsm_UndoCheckOut = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_CheckInFile(pageUrl As String, comment As String, CheckinType As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "CheckInFile"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "pageUrl", pageUrl)
    Call AddAttributes(attrs, "comment", comment)
    Call AddAttributes(attrs, "CheckinType", CheckinType)

    body = Soap_Body(method, attrs)
    Set wsm_CheckInFile = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListContentTypes(listName As String, contentTypeId As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListContentTypes"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)

    body = Soap_Body(method, attrs)
    Set wsm_GetListContentTypes = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListContentTypesAndProperties(listName As String, contentTypeId As String, propertyPrefix As String, includeWebProperties As Boolean, includeWebPropertiesSpecified As Boolean) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListContentTypesAndProperties"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)
    Call AddAttributes(attrs, "propertyPrefix", propertyPrefix)
    Call AddAttributes(attrs, "includeWebProperties", includeWebProperties)
    Call AddAttributes(attrs, "includeWebPropertiesSpecified", includeWebPropertiesSpecified)

    body = Soap_Body(method, attrs)
    Set wsm_GetListContentTypesAndProperties = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListContentType(listName As String, contentTypeId As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListContentType"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)

    body = Soap_Body(method, attrs)
    Set wsm_GetListContentType = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_CreateContentType(listName As String, displayName As String, parentType As String, fields As MSXML2.IXMLDOMNodeList, contentTypeProperties As MSXML2.IXMLDOMNodeList, addToView As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "CreateContentType"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "displayName", displayName)
    Call AddAttributes(attrs, "parentType", parentType)
    Call AddAttributes(attrs, "fields", fields)
    Call AddAttributes(attrs, "contentTypeProperties", contentTypeProperties)
    Call AddAttributes(attrs, "addToView", addToView)

    body = Soap_Body(method, attrs)
    Set wsm_CreateContentType = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateContentType(listName As String, contentTypeId As String, contentTypeProperties As MSXML2.IXMLDOMNodeList, newFields As MSXML2.IXMLDOMNodeList, updateFields As MSXML2.IXMLDOMNodeList, deleteFields As MSXML2.IXMLDOMNodeList, addToView As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateContentType"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)
    Call AddAttributes(attrs, "contentTypeProperties", contentTypeProperties)
    Call AddAttributes(attrs, "newFields", newFields)
    Call AddAttributes(attrs, "updateFields", updateFields)
    Call AddAttributes(attrs, "deleteFields", deleteFields)
    Call AddAttributes(attrs, "addToView", addToView)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateContentType = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_DeleteContentType(listName As String, contentTypeId As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "DeleteContentType"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)

    body = Soap_Body(method, attrs)
    Set wsm_DeleteContentType = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateContentTypeXmlDocument(listName As String, contentTypeId As String, newDocument As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateContentTypeXmlDocument"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)
    Call AddAttributes(attrs, "newDocument", newDocument)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateContentTypeXmlDocument = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateContentTypesXmlDocument(listName As String, newDocument As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateContentTypesXmlDocument"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "newDocument", newDocument)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateContentTypesXmlDocument = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_DeleteContentTypeXmlDocument(listName As String, contentTypeId As String, documentUri As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "DeleteContentTypeXmlDocument"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)
    Call AddAttributes(attrs, "documentUri", documentUri)

    body = Soap_Body(method, attrs)
    Set wsm_DeleteContentTypeXmlDocument = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_ApplyContentTypeToList(webUrl As String, contentTypeId As String, listName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "ApplyContentTypeToList"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "webUrl", webUrl)
    Call AddAttributes(attrs, "contentTypeId", contentTypeId)
    Call AddAttributes(attrs, "listName", listName)

    body = Soap_Body(method, attrs)
    Set wsm_ApplyContentTypeToList = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetList(listName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetList"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)

    body = Soap_Body(method, attrs)
    Set wsm_GetList = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListAndView(listName As String, viewName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListAndView"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "viewName", viewName)

    body = Soap_Body(method, attrs)
    Set wsm_GetListAndView = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_DeleteList(listName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "DeleteList"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)

    body = Soap_Body(method, attrs)
    Set wsm_DeleteList = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_AddList(listName As String, description As String, templateID As Integer) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddList"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "description", description)
    Call AddAttributes(attrs, "templateID", templateID)

    body = Soap_Body(method, attrs)
    Set wsm_AddList = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_AddListFromFeature(listName As String, description As String, featureID As String, templateID As Integer) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddListFromFeature"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "description", description)
    Call AddAttributes(attrs, "featureID", featureID)
    Call AddAttributes(attrs, "templateID", templateID)

    body = Soap_Body(method, attrs)
    Set wsm_AddListFromFeature = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateList(listName As String, listProperties As MSXML2.IXMLDOMNodeList, newFields As MSXML2.IXMLDOMNodeList, updateFields As MSXML2.IXMLDOMNodeList, deleteFields As MSXML2.IXMLDOMNodeList, listVersion As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateList"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "listProperties", listProperties)
    Call AddAttributes(attrs, "newFields", newFields)
    Call AddAttributes(attrs, "updateFields", updateFields)
    Call AddAttributes(attrs, "deleteFields", deleteFields)
    Call AddAttributes(attrs, "listVersion", listVersion)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateList = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListCollection() As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListCollection"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    body = Soap_Body(method, attrs)
    Set wsm_GetListCollection = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListItems(listName As String, viewName As String, Query As MSXML2.IXMLDOMNodeList, viewFields As MSXML2.IXMLDOMNodeList, rowLimit As String, queryOptions As MSXML2.IXMLDOMNodeList, webID As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListItems"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "viewName", viewName)
    Call AddAttributes(attrs, "query", Query)
    Call AddAttributes(attrs, "viewFields", viewFields)
    Call AddAttributes(attrs, "rowLimit", rowLimit)
    Call AddAttributes(attrs, "queryOptions", queryOptions)
    Call AddAttributes(attrs, "webID", webID)

    body = Soap_Body(method, attrs)
    Set wsm_GetListItems = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListItemChanges(listName As String, viewFields As MSXML2.IXMLDOMNodeList, since As String, contains As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListItemChanges"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "viewFields", viewFields)
    Call AddAttributes(attrs, "since", since)
    Call AddAttributes(attrs, "contains", contains)

    body = Soap_Body(method, attrs)
    Set wsm_GetListItemChanges = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListItemChangesWithKnowledge(listName As String, viewName As String, Query As MSXML2.IXMLDOMNodeList, viewFields As MSXML2.IXMLDOMNodeList, rowLimit As String, queryOptions As MSXML2.IXMLDOMNodeList, syncScope As String, knowledge As MSXML2.IXMLDOMNodeList, contains As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListItemChangesWithKnowledge"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "viewName", viewName)
    Call AddAttributes(attrs, "query", Query)
    Call AddAttributes(attrs, "viewFields", viewFields)
    Call AddAttributes(attrs, "rowLimit", rowLimit)
    Call AddAttributes(attrs, "queryOptions", queryOptions)
    Call AddAttributes(attrs, "syncScope", syncScope)
    Call AddAttributes(attrs, "knowledge", knowledge)
    Call AddAttributes(attrs, "contains", contains)

    body = Soap_Body(method, attrs)
    Set wsm_GetListItemChangesWithKnowledge = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetListItemChangesSinceToken(listName As String, viewName As String, Query As MSXML2.IXMLDOMNodeList, viewFields As MSXML2.IXMLDOMNodeList, rowLimit As String, queryOptions As MSXML2.IXMLDOMNodeList, changeToken As String, contains As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetListItemChangesSinceToken"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "viewName", viewName)
    Call AddAttributes(attrs, "query", Query)
    Call AddAttributes(attrs, "viewFields", viewFields)
    Call AddAttributes(attrs, "rowLimit", rowLimit)
    Call AddAttributes(attrs, "queryOptions", queryOptions)
    Call AddAttributes(attrs, "changeToken", changeToken)
    Call AddAttributes(attrs, "contains", contains)

    body = Soap_Body(method, attrs)
    Set wsm_GetListItemChangesSinceToken = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateListItems(listName As String, updates As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateListItems"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "updates", updates)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateListItems = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_UpdateListItemsWithKnowledge(listName As String, updates As MSXML2.IXMLDOMNodeList, syncScope As String, knowledge As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateListItemsWithKnowledge"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "updates", updates)
    Call AddAttributes(attrs, "syncScope", syncScope)
    Call AddAttributes(attrs, "knowledge", knowledge)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateListItemsWithKnowledge = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_AddDiscussionBoardItem(listName As String, message As Variant) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddDiscussionBoardItem"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "listName", listName)
    Call AddAttributes(attrs, "message", message)

    body = Soap_Body(method, attrs)
    Set wsm_AddDiscussionBoardItem = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_AddWikiPage(strlistName As String, listRelPageUrl As String, wikiContent As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddWikiPage"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strListName", strlistName)
    Call AddAttributes(attrs, "listRelPageUrl", listRelPageUrl)
    Call AddAttributes(attrs, "wikiContent", wikiContent)

    body = Soap_Body(method, attrs)
    Set wsm_AddWikiPage = Request("/_vti_bin/Lists.asmx", method, body)

End Function

Public Function wsm_GetView(strlistName As String, strviewName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetView"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)

    body = Soap_Body(method, attrs)
    Set wsm_GetView = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_GetViewHtml(strlistName As String, strviewName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetViewHtml"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)

    body = Soap_Body(method, attrs)
    Set wsm_GetViewHtml = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_DeleteView(strlistName As String, strviewName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "DeleteView"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)

    body = Soap_Body(method, attrs)
    Set wsm_DeleteView = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_AddView(strlistName As String, strviewName As String, xmlviewFields As MSXML2.IXMLDOMNodeList, xmlquery As MSXML2.IXMLDOMNodeList, xmlrowLimit As MSXML2.IXMLDOMNodeList, strtype As String, bmakeViewDefault As Boolean) As MSXML2.IXMLDOMNodeList

    Const method As String = "AddView"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)
    Call AddAttributes(attrs, "xmlviewFields", xmlviewFields)
    Call AddAttributes(attrs, "xmlquery", xmlquery)
    Call AddAttributes(attrs, "xmlrowLimit", xmlrowLimit)
    Call AddAttributes(attrs, "strtype", strtype)
    Call AddAttributes(attrs, "bmakeViewDefault", bmakeViewDefault)

    body = Soap_Body(method, attrs)
    Set wsm_AddView = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_GetViewCollection(strlistName As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "GetViewCollection"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)

    body = Soap_Body(method, attrs)
    Set wsm_GetViewCollection = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_UpdateView(strlistName As String, strviewName As String, xmlviewProperties As MSXML2.IXMLDOMNodeList, xmlquery As MSXML2.IXMLDOMNodeList, xmlviewFields As MSXML2.IXMLDOMNodeList, xmlaggregations As MSXML2.IXMLDOMNodeList, xmlformats As MSXML2.IXMLDOMNodeList, xmlrowLimit As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateView"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)
    Call AddAttributes(attrs, "xmlviewProperties", xmlviewProperties)
    Call AddAttributes(attrs, "xmlquery", xmlquery)
    Call AddAttributes(attrs, "xmlviewFields", xmlviewFields)
    Call AddAttributes(attrs, "xmlaggregations", xmlaggregations)
    Call AddAttributes(attrs, "xmlformats", xmlformats)
    Call AddAttributes(attrs, "xmlrowLimit", xmlrowLimit)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateView = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_UpdateViewHtml(strlistName As String, strviewName As String, xmlviewProperties As MSXML2.IXMLDOMNodeList, xmltoolbar As MSXML2.IXMLDOMNodeList, xmlviewHeader As MSXML2.IXMLDOMNodeList, xmlviewBody As MSXML2.IXMLDOMNodeList, xmlviewFooter As MSXML2.IXMLDOMNodeList, xmlviewEmpty As MSXML2.IXMLDOMNodeList, xmlrowLimitExceeded As MSXML2.IXMLDOMNodeList, xmlquery As MSXML2.IXMLDOMNodeList, xmlviewFields As MSXML2.IXMLDOMNodeList, xmlaggregations As MSXML2.IXMLDOMNodeList, xmlformats As MSXML2.IXMLDOMNodeList, xmlrowLimit As MSXML2.IXMLDOMNodeList) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateViewHtml"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)
    Call AddAttributes(attrs, "xmlviewProperties", xmlviewProperties)
    Call AddAttributes(attrs, "xmltoolbar", xmltoolbar)
    Call AddAttributes(attrs, "xmlviewHeader", xmlviewHeader)
    Call AddAttributes(attrs, "xmlviewBody", xmlviewBody)
    Call AddAttributes(attrs, "xmlviewFooter", xmlviewFooter)
    Call AddAttributes(attrs, "xmlviewEmpty", xmlviewEmpty)
    Call AddAttributes(attrs, "xmlrowLimitExceeded", xmlrowLimitExceeded)
    Call AddAttributes(attrs, "xmlquery", xmlquery)
    Call AddAttributes(attrs, "xmlviewFields", xmlviewFields)
    Call AddAttributes(attrs, "xmlaggregations", xmlaggregations)
    Call AddAttributes(attrs, "xmlformats", xmlformats)
    Call AddAttributes(attrs, "xmlrowLimit", xmlrowLimit)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateViewHtml = Request("/_vti_bin/Views.asmx", method, body)

End Function

Public Function wsm_UpdateViewHtml2(strlistName As String, strviewName As String, xmlviewProperties As MSXML2.IXMLDOMNodeList, xmltoolbar As MSXML2.IXMLDOMNodeList, xmlviewHeader As MSXML2.IXMLDOMNodeList, xmlviewBody As MSXML2.IXMLDOMNodeList, xmlviewFooter As MSXML2.IXMLDOMNodeList, xmlviewEmpty As MSXML2.IXMLDOMNodeList, xmlrowLimitExceeded As MSXML2.IXMLDOMNodeList, xmlquery As MSXML2.IXMLDOMNodeList, xmlviewFields As MSXML2.IXMLDOMNodeList, xmlaggregations As MSXML2.IXMLDOMNodeList, xmlformats As MSXML2.IXMLDOMNodeList, xmlrowLimit As MSXML2.IXMLDOMNodeList, stropenApplicationExtension As String) As MSXML2.IXMLDOMNodeList

    Const method As String = "UpdateViewHtml2"

    Dim attrs As Collection
    Set attrs = New Collection
    Dim body As String

    Call AddAttributes(attrs, "strlistName", strlistName)
    Call AddAttributes(attrs, "strviewName", strviewName)
    Call AddAttributes(attrs, "xmlviewProperties", xmlviewProperties)
    Call AddAttributes(attrs, "xmltoolbar", xmltoolbar)
    Call AddAttributes(attrs, "xmlviewHeader", xmlviewHeader)
    Call AddAttributes(attrs, "xmlviewBody", xmlviewBody)
    Call AddAttributes(attrs, "xmlviewFooter", xmlviewFooter)
    Call AddAttributes(attrs, "xmlviewEmpty", xmlviewEmpty)
    Call AddAttributes(attrs, "xmlrowLimitExceeded", xmlrowLimitExceeded)
    Call AddAttributes(attrs, "xmlquery", xmlquery)
    Call AddAttributes(attrs, "xmlviewFields", xmlviewFields)
    Call AddAttributes(attrs, "xmlaggregations", xmlaggregations)
    Call AddAttributes(attrs, "xmlformats", xmlformats)
    Call AddAttributes(attrs, "xmlrowLimit", xmlrowLimit)
    Call AddAttributes(attrs, "stropenApplicationExtension", stropenApplicationExtension)

    body = Soap_Body(method, attrs)
    Set wsm_UpdateViewHtml2 = Request("/_vti_bin/Views.asmx", method, body)

End Function



