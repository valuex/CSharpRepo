using System;
using System.Xml;
using System.Xml.Linq;
using Append2OneNote;
using Microsoft.Office.Interop.OneNote;


class Program
{
    static void Main()
    {
        Microsoft.Office.Interop.OneNote.Application onenoteApp = new Microsoft.Office.Interop.OneNote.Application();


        string pageId = "{3ACAE6C6-53C5-0C4F-060D-3007A1428E63}{1}{E19561840189956234832620129833010512773022001}";

        onenoteApp.GetPageContent(pageId, out var CurrentPageXml);

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(CurrentPageXml);

        XmlNode root = xmlDoc.DocumentElement;
        var LastOutline = root.LastChild;
        var LastOEChildren = LastOutline.LastChild;
        var LastOE = LastOEChildren.LastChild;

        var pageContent = "Your text content1";
        string newXML = $@"
                <one:OE alignment=""left"" quickStyleIndex=""1"" selected=""partial"" xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:T selected=""all""><![CDATA[{pageContent}]]></one:T>
                </one:OE>";

        XmlDocumentFragment xfrag = xmlDoc.CreateDocumentFragment();
        xfrag.InnerXml = newXML;

 
        LastOEChildren.AppendChild(xfrag);

        onenoteApp.UpdatePageContent(xmlDoc.InnerXml);
    }
}
