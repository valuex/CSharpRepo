using System;
using Microsoft.Office.Interop.OneNote;
using static System.Net.Mime.MediaTypeNames;

class Program
{
    static void Main()
    {
        // Initialize OneNote application
        Microsoft.Office.Interop.OneNote.Application onenoteApp = new Microsoft.Office.Interop.OneNote.Application();
        // Create a new page
        var notebookId = "{6166EEB3-9FC7-4534-A8BA-BD4E41FC2946}{1}{B0}";
        var sectionId = "{3ACAE6C6-53C5-0C4F-060D-3007A1428E63}{1}{B0}";
        var pageTitle = "Your Page Title";
        var pageContent = "Your text content";

        string newPageId;
        //onenoteApp.CreateNewPage(sectionId, pageTitle, out newPageId);
        onenoteApp.CreateNewPage(sectionId, out var pId, NewPageStyle.npsBlankPageNoTitle);

        // Get the content XML for the page
        var contentXml = $@"
            <one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote"" ID=""{pId}"" dateTime=""2024-02-24T02:14:22.000Z"" lastModifiedTime=""2024-02-24T02:14:56.000Z"" pageLevel=""1"" isCurrentlyViewed=""true"" selected=""partial"" lang=""zh-CN"">
            <one:QuickStyleDef index=""0"" name=""PageTitle"" fontColor=""automatic"" highlightColor=""automatic"" font=""Calibri Light"" fontSize=""20.0"" spaceBefore=""0.0"" spaceAfter=""0.0"" />
            <one:QuickStyleDef index=""1"" name=""p"" fontColor=""automatic"" highlightColor=""automatic"" font=""Calibri"" fontSize=""11.0"" spaceBefore=""0.0"" spaceAfter=""0.0"" />
            <one:PageSettings RTL=""false"" color=""automatic"">
            <one:PageSize>
                <one:Automatic />
            </one:PageSize>
            <one:RuleLines visible=""false"" />
            </one:PageSettings>
            <one:Title lang=""zh-CN"">
            <one:OE alignment=""left"" quickStyleIndex=""0"">
                <one:T><![CDATA[{pageTitle}]]></one:T>
            </one:OE>
            </one:Title>
            <one:Outline selected=""partial"">
            <one:Position x=""54"" y=""86"" z=""0"" />
            <one:Size width=""72.0"" height=""13.42771339416504"" />
            <one:OEChildren selected=""partial"">
                <one:OE alignment=""left"" quickStyleIndex=""1"" selected=""partial"">
                <one:T selected=""all""><![CDATA[{pageContent}]]></one:T>
                </one:OE>
            </one:OEChildren>
            </one:Outline>
        </one:Page>";

        // Update the content of the page
        onenoteApp.UpdatePageContent(contentXml);

        Console.WriteLine("Text sent to OneNote Desktop successfully!");
    }
}
