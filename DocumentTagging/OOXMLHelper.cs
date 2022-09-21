﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.IO;
using We = DocumentFormat.OpenXml.Office2013.WebExtension;
using Wetp = DocumentFormat.OpenXml.Office2013.WebExtentionPane;
using System;
using System.Configuration;

namespace DocumentTagging
{
    public static class OOXMLHelper
    {
        public static void ValidateOfficeFile(string fileType, MemoryStream memoryStream)
        {
            // Validate the file by trying to open it.
            // Exception is thrown if it is not valid.
            switch (fileType)
            {
                // Each Office file type has its own *Document class in the OOXML SDK.
                case "Excel":
                    using (var spreadsheet = SpreadsheetDocument.Open(memoryStream, true)) { }
                    break;
                case "Word":
                    using (var doc = WordprocessingDocument.Open(memoryStream, true)) { }
                    break;
                case "PowerPoint":
                    using (var slidedeck = PresentationDocument.Open(memoryStream, true)) { }
                    break;
                default:
                    throw new Exception("Only Excel, Word, and PowerPoint files can be validated.");
            }
        }

        /*
         * Except for code enclosed in "CUSTOM MODIFICATION", all code and comments below this point 
         * were generated by the Open XML SDK 2.5 Productivity Tool which you can get from here:
         * https://www.microsoft.com/en-us/download/details.aspx?id=30425
         */

        // Adds child parts and generates content of the specified part.
        public static void CreateWebExTaskpanesPart(WebExTaskpanesPart part, int id)
        {
            WebExtensionPart webExtensionPart1 = part.AddNewPart<WebExtensionPart>("rId" + id);
            GenerateWebExtensionPart1Content(webExtensionPart1);

            GeneratePartContent(part, id);
        }

        // Generates content of webExtensionPart1.
        private static void GenerateWebExtensionPart1Content(WebExtensionPart webExtensionPart1)
        {
            We.WebExtension webExtension1 = new We.WebExtension() { Id = "{" + Guid.NewGuid() + "}" };
            webExtension1.AddNamespaceDeclaration("we", "http://schemas.microsoft.com/office/webextensions/webextension/2010/11");
            We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { 
                Id = ConfigurationManager.AppSettings.Get("AddinId"),
                Version = ConfigurationManager.AppSettings.Get("AddinVersion"),
                Store = ConfigurationManager.AppSettings.Get("AddinStore"), 
                StoreType = ConfigurationManager.AppSettings.Get("AddinStoreType")
            };
            We.WebExtensionReferenceList webExtensionReferenceList1 = new We.WebExtensionReferenceList();

            We.WebExtensionPropertyBag webExtensionPropertyBag1 = new We.WebExtensionPropertyBag();

            // Add the property that makes the taskpane visible.
            We.WebExtensionProperty webExtensionProperty1 = new We.WebExtensionProperty() { Name = "Office.AutoShowTaskpaneWithDocument", Value = "true" };
            webExtensionPropertyBag1.Append(webExtensionProperty1);

            We.WebExtensionBindingList webExtensionBindingList1 = new We.WebExtensionBindingList();

            We.Snapshot snapshot1 = new We.Snapshot();
            snapshot1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            webExtension1.Append(webExtensionStoreReference1);
            webExtension1.Append(webExtensionReferenceList1);
            webExtension1.Append(webExtensionPropertyBag1);
            webExtension1.Append(webExtensionBindingList1);
            webExtension1.Append(snapshot1);

            webExtensionPart1.WebExtension = webExtension1;
        }

        // Generates content of part.
        private static void GeneratePartContent(WebExTaskpanesPart part, int id)
        {

            Wetp.WebExtensionTaskpane webExtensionTaskpane1 = new Wetp.WebExtensionTaskpane() { DockState = "right", Visibility = true, Width = 350D, Row = (UInt32Value)4U };

            Wetp.WebExtensionPartReference webExtensionPartReference1 = new Wetp.WebExtensionPartReference() { Id = "rId" + id };
            webExtensionPartReference1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            webExtensionTaskpane1.Append(webExtensionPartReference1);

            if (part.Taskpanes != null)
			{
                var taskpanes = part.Taskpanes;

                taskpanes.Append(webExtensionTaskpane1);

            } else
			{
                Wetp.Taskpanes taskpanes1 = new Wetp.Taskpanes();
                taskpanes1.AddNamespaceDeclaration("wetp", "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11");

                taskpanes1.Append(webExtensionTaskpane1);

                part.Taskpanes = taskpanes1;

            }
        }
    }
}