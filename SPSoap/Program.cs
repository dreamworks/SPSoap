using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1 {

    public class member_attr {
        public string name;
        public string type;
    }

    class Program {
        static void Main(string[] args) {
            

            List<string> methods = new List<string>() {
                 "AddAttachment",
                 "AddDiscussionBoardItem",
                 "AddList",
                 "AddListFromFeature",
                 "AddWikiPage",
                 "ApplyContentTypeToList",
                 "CheckInFile",
                 "CheckOutFile",
                 "CreateContentType",
                 "DeleteAttachment",
                 "DeleteContentType",
                 "DeleteContentTypeXmlDocument",
                 "DeleteList",
                 "GetAttachmentCollection",
                 "GetList",
                 "GetListAndView",
                 "GetListCollection",
                 "GetListContentType",
                 "GetListContentTypes",
                 "GetListContentTypesAndProperties",
                 "GetListItemChanges",
                 "GetListItemChangesSinceToken",
                 "GetListItemChangesWithKnowledge",
                 "GetListItems",
                 "GetVersionCollection",
                 "UndoCheckOut",
                 "UpdateContentType",
                 "UpdateContentTypeXmlDocument",
                 "UpdateContentTypesXmlDocument",
                 "UpdateList",
                 "UpdateListItems",
                 "UpdateListItemsWithKnowledge" 
            };


            List<string> SOAPTypes = new List<string>();


            MethodInfo[] methodInfos = typeof(SPSOAP.Lists).GetMethods(BindingFlags.Public | BindingFlags.Instance);

            foreach (var m in methodInfos) {
                if (methods.Contains(m.Name)) {

                    List<member_attr> attrs = new List<member_attr>();
                    for (int i = 0; i < m.GetParameters().Length; i++) {
                        string fulltype  = m.GetParameters().GetValue(i).ToString();

                        string[] data = fulltype.Split(' ');

                        string type = data[0].Trim();
                        string name = data[1];

                         if (!SOAPTypes.Contains(type)) {
                            SOAPTypes.Add(type);
                        }

                        switch (type) {

                            case "Boolean":
                                type = "Boolean";
                                break;

                            case "System.Xml.XmlNode":
                                type = "MSXML2.IXMLDOMNodeList";
                                break;

                            case "Int32":
                                type = "Integer";
                                break;

                            case "Byte[]":
                                type = "Variant";
                                break;
                            default:
                                type = "String";
                                break;
                        }

                        member_attr attr = new member_attr();
                        attr.name = name;
                        attr.type = type;
                        
                        attrs.Add(attr);



                    }

                    string ret_Type = m.ReturnParameter.ToString().Trim();

                    if (!SOAPTypes.Contains(ret_Type)) {
                        SOAPTypes.Add(ret_Type);
                    }

                    switch (ret_Type) {
                        /*
                        case "Boolean":
                            ret_Type = "Boolean";
                            break;

                        case "System.Xml.XmlNode":
                            ret_Type = "MSXML2.IXMLDOMNodeList";
                            break;

                        case "Byte[]":
                            ret_Type = "Variant";
                            break;
                        
                        default:
                            ret_Type = "String";
                            break;
                        */
                        default:
                            ret_Type = "MSXML2.IXMLDOMNodeList";
                            break;
                    }


                    string header = "";
                    if (attrs.Count > 0) {
                        foreach (member_attr attr in attrs) {
                            header += attr.name + " As " + attr.type + ", ";
                        }
                        header = header.Substring(0,header.Length - 2);

                    }


                    Console.WriteLine("Public Function wsm_" + m.Name + "(" + header + ") As " + ret_Type);


                    Console.WriteLine();
                    Console.WriteLine("    Const method As String = \"" + m.Name + "\"");
                    Console.WriteLine();
                    Console.WriteLine("    Dim attrs As Collection");
                    Console.WriteLine("    Set attrs = New Collection");
                    Console.WriteLine("    Dim body As String");
                    if (attrs.Count > 0) {
                        Console.WriteLine();
                    }
                    foreach (var attr in attrs) {
                        Console.WriteLine("    Call AddAttributes(attrs, \"" + attr.name + "\", " + attr.name + ")");
                    }
                    Console.WriteLine();
                    Console.WriteLine("    body = Soap_Body(method, attrs)");
                    Console.WriteLine("    Set wsm_"+m.Name + " = Request(method, body)");
                    Console.WriteLine();
                    Console.WriteLine("End Function");
                    Console.WriteLine();
                    
                }

            }



            foreach (var type in SOAPTypes) {
                Console.WriteLine("Type: " + type);
            }
            
            
        }
    }
}
