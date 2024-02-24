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
        string imagePath = "C:\\Users\\wei_x\\OneDrive\\Desktop\\GetTitle\\doc\\page_1.jpg";
        // Convert the image to base64
        byte[] imageBytes = File.ReadAllBytes(imagePath);
        string base64Image = Convert.ToBase64String(imageBytes);
        string ImageDescription = "page_1";
        string ImageExt = "jpg";
        //<one:Data format=""{ImageExt}"" alt=""{ImageDescription}"" data=""{base64Image}""></one:Data>
        string newXML = $@"
                <one:OE alignment=""left"" quickStyleIndex=""1"" xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"">
                <one:Image>
                <one:Data>{base64Image}</one:Data>
                </one:Image>
                </one:OE>";

        onenoteApp.GetPageContent(pageId, out var CurrentPageXml);

        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(CurrentPageXml);

        XmlNode root = xmlDoc.DocumentElement;
        var LastOutline = root.LastChild;
        var LastOEChildren = LastOutline.LastChild;
        var LastOE = LastOEChildren.LastChild;

        XmlDocumentFragment xfrag = xmlDoc.CreateDocumentFragment();
        xfrag.InnerXml = newXML; 
        LastOEChildren.AppendChild(xfrag);
        onenoteApp.UpdatePageContent(xmlDoc.InnerXml);
    }
}
