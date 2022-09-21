using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace DocumentTagging
{
    public static class AddinEmbedder
    {
        // Embeds the add-in into a file of the specified type.
        public static void EmbedAddin(string fileType, Stream memoryStream)
        {
            // Each Office file type has its own *Document class in the OOXML SDK.
            switch (fileType)
            {
                case "Excel":
                    using (var spreadsheet = SpreadsheetDocument.Open(memoryStream, true))
                    {
                        //spreadsheet.DeletePart(spreadsheet.WebExTaskpanesPart);
                        if (spreadsheet.WebExTaskpanesPart != null)
						{
                            var webExTaskpanesPart = spreadsheet.WebExTaskpanesPart;
                            int count = 1;
                            foreach (var WebExtensionPart in webExTaskpanesPart.WebExtensionParts)
                            {
                                count++;
                            }

                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, count);
                        } else
						{
                            var webExTaskpanesPart = spreadsheet.AddWebExTaskpanesPart();
                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, 1);
                        }
                    }
                    break;
                case "Word":
                    using (var document = WordprocessingDocument.Open(memoryStream, true))
                    {
                        if (document.WebExTaskpanesPart != null)
						{
                            var webExTaskpanesPart = document.WebExTaskpanesPart;
                            int count = 1;
                            foreach (var WebExtensionPart in webExTaskpanesPart.WebExtensionParts)
							{
                                count++;
							}

                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, count);

                        } else
						{
                            var webExTaskpanesPart = document.AddWebExTaskpanesPart();

                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, 1);
                        }
                    }
                    break;
                case "PowerPoint":
                    using (var slidedeck = PresentationDocument.Open(memoryStream, true))
                    {
                        if (slidedeck.WebExTaskpanesPart != null)
                        {
                            var webExTaskpanesPart = slidedeck.WebExTaskpanesPart;
                            int count = 1;
                            foreach (var WebExtensionPart in webExTaskpanesPart.WebExtensionParts)
                            {
                                count++;
                            }

                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, count);

                        }
                        else
                        {
                            var webExTaskpanesPart = slidedeck.AddWebExTaskpanesPart();
                            OOXMLHelper.CreateWebExTaskpanesPart(webExTaskpanesPart, 1);
                        }
                    }
                    break;
                default:
                    break;
                    //throw new Exception("Invalid File Type");
            }
        }
    }
}