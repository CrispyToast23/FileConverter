using System;
using System.IO.Packaging;
using System.Xml.Linq;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Xml;

namespace FileConverter
{
    public class WordToPdfConverter
    {
        public void ConvertWordToPdf(string wordFilePath, string pdfFilePath)
        {
            // Open the Word document as a ZIP package
            using (Package wordPackage = Package.Open(wordFilePath, FileMode.Open, FileAccess.Read))
            {
                // Get the document.xml file inside the package
                PackagePart documentPart = wordPackage.GetPart(new Uri("/word/document.xml", UriKind.Relative));
                XDocument documentXml = XDocument.Load(XmlReader.Create(documentPart.GetStream()));

                // Extract text from paragraphs and handle line breaks and paragraph formatting
                string content = ExtractContentFromWord(documentXml);

                // Now create the PDF
                PdfGenerator pdfGenerator = new PdfGenerator();
                pdfGenerator.CreatePdf(pdfFilePath, content, documentXml);
            }
        }

        private string ExtractContentFromWord(XDocument documentXml)
        {
            string content = "";
            // Loop through each paragraph in the Word document
            foreach (var paragraph in documentXml.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
            {
                // Extract text from each run within the paragraph
                foreach (var run in paragraph.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r"))
                {
                    var textElement = run.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t");
                    if (textElement != null)
                    {
                        // Append text with proper handling for spaces and line breaks
                        content += textElement.Value;
                    }
                }
                content += "\n"; // Add newline after each paragraph
            }
            return content;
        }
    }

    public class PdfGenerator
    {
        public void CreatePdf(string pdfFilePath, string content, XDocument documentXml)
        {
            PdfDocument document = new PdfDocument();
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Arial", 12);
            double xPosition = 40; // Default left alignment
            double yPosition = 40; // Starting position

            // Split content into lines based on newline character
            string[] lines = content.Split(new[] { "\n" }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                // Get paragraph formatting from Word document
                var paragraph = GetParagraphFromContent(documentXml, line);

                if (paragraph == null)
                    continue; // Skip if the paragraph is null

                // Determine text color
                XBrush textColor = GetTextColor(paragraph);

                // Set alignment based on Word document's paragraph alignment
                var alignment = GetParagraphAlignment(paragraph);
                switch (alignment)
                {
                    case "center":
                        xPosition = page.Width / 2 - gfx.MeasureString(line, font).Width / 2; // Center the text
                        break;
                    case "right":
                        xPosition = page.Width - 40 - gfx.MeasureString(line, font).Width; // Align right
                        break;
                    default:
                        xPosition = 40; // Default left alignment
                        break;
                }

                // Draw the text with the determined position and color
                gfx.DrawString(line, font, textColor, xPosition, yPosition);
                yPosition += font.GetHeight() + 5; // Use GetHeight() without parameter to determine line height and add a small margin

                // Handle page overflow (if the content goes beyond the page)
                if (yPosition > page.Height - 40)
                {
                    page = document.AddPage();
                    gfx = XGraphics.FromPdfPage(page);
                    yPosition = 40; // Reset position for the new page
                }
            }

            // Save the PDF to a file
            document.Save(pdfFilePath);
        }

        private XBrush GetTextColor(XElement paragraph)
        {
            // Check for font color in the Word document's run element and apply color in PDF
            var colorElement = paragraph.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color").FirstOrDefault();
            if (colorElement != null)
            {
                string color = colorElement.Attribute("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")?.Value;

                // Ensure the color is in the correct format and length
                if (!string.IsNullOrEmpty(color) && color.Length == 7 && color[0] == '#')
                {
                    try
                    {
                        // Assuming color is in hex format #RRGGBB
                        int r = int.Parse(color.Substring(1, 2), System.Globalization.NumberStyles.HexNumber);
                        int g = int.Parse(color.Substring(3, 2), System.Globalization.NumberStyles.HexNumber);
                        int b = int.Parse(color.Substring(5, 2), System.Globalization.NumberStyles.HexNumber);

                        return new XSolidBrush(XColor.FromArgb(r, g, b));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error parsing color: {ex.Message}");
                    }
                }
            }
            return XBrushes.Black; // Default to black if no color found or invalid format
        }

        private string GetParagraphAlignment(XElement paragraph)
        {
            // Check for paragraph alignment in the Word document
            var alignmentElement = paragraph.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc").FirstOrDefault();
            if (alignmentElement != null)
            {
                string alignmentValue = alignmentElement.Attribute("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")?.Value;
                switch (alignmentValue)
                {
                    case "center":
                        return "center";
                    case "right":
                        return "right";
                    default:
                        return "left"; // Default alignment
                }
            }
            return "left"; // Default to left if no alignment found
        }

        private XElement GetParagraphFromContent(XDocument documentXml, string content)
        {
            // This function can be further optimized to fetch the paragraph based on the content
            return documentXml.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
                .FirstOrDefault(p => p.Descendants("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
                .Any(t => t.Value.Contains(content)));
        }
    }
}
