using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GeneratedPDFFile.Services;

public class HtmlToDocxFileService
{
    private string GetContentFromTag(string tag, ref string body)
    {
        string pattern = $@"(?:<{tag}>(?<content>(?:.*?\r?\n?)*)<\/{tag}>)+";
        RegexOptions options = RegexOptions.Multiline;
        Regex expression = new Regex(pattern, options);
        Match match = expression.Match(body);
        if (match.Success)
        {
            // Remove tag from body
            Regex regex = new Regex(pattern, options);
            body = regex.Replace(body, "");

            return match.Groups["content"].Value;
        }
        else
            return null;
    }
    private IList<OpenXmlCompositeElement> ConvertHtmlToOpenXml(string input)
    {
        using (MemoryStream generatedDocument = new MemoryStream())
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = package.MainDocumentPart;
                if (mainPart == null)
                {
                    mainPart = package.AddMainDocumentPart();
                    new Document(new Body()).Save(mainPart);
                }

                HtmlToOpenXml.HtmlConverter converter = new HtmlToOpenXml.HtmlConverter(mainPart);
                return converter.Parse(input);
            }
        }
    }
    
    private void ApplyFooter(WordprocessingDocument doc, string input)
    {
        if (!string.IsNullOrWhiteSpace(input))
        {
            MainDocumentPart mainDocPart = doc.MainDocumentPart;
            FooterPart footerPart = mainDocPart.AddNewPart<FooterPart>("r98");

            Footer footer = new Footer();
            Paragraph paragraph = new Paragraph() { };
            Run run = new Run();
            run.Append(ConvertHtmlToOpenXml(input));
            paragraph.Append(run);
            footer.Append(paragraph);
            footerPart.Footer = footer;

            SectionProperties sectionProperties = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProperties == null)
            {
                sectionProperties = new SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties);
            }

            FooterReference footerReference = new FooterReference() { Type =HeaderFooterValues.Default, Id = "r98" };
            sectionProperties.InsertAt(footerReference, 0);
        }
    }
    private void ApplyHeader(WordprocessingDocument doc, string input)
    {
        if (!string.IsNullOrWhiteSpace(input))
        {
            MainDocumentPart mainDocPart = doc.MainDocumentPart;
            HeaderPart headerPart = mainDocPart.AddNewPart<HeaderPart>("r97");

            Header header = new Header();
            Paragraph paragraph = new Paragraph() { };
            Run run = new Run();
            run.Append(ConvertHtmlToOpenXml(input));
            paragraph.Append(run);
            header.Append(paragraph);
            headerPart.Header = header;

            SectionProperties sectionProperties = mainDocPart.Document.Body.Descendants<SectionProperties>().FirstOrDefault();
            if (sectionProperties == null)
            {
                sectionProperties = new SectionProperties() { };
                mainDocPart.Document.Body.Append(sectionProperties);
            }

            HeaderReference headerReference = new HeaderReference() { Type = DocumentFormat.OpenXml.Wordprocessing.HeaderFooterValues.Default, Id = "r97" };
            sectionProperties.InsertAt(headerReference, 0);
        }
    }
}