using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace GeneratedPDFFile.Services.Interfaces;

public interface IHtmlToDocxFileService
{
    string GetContentFromTag(string tag, ref string body);
    IList<OpenXmlCompositeElement> ConvertHtmlToOpenXml(string input);
    void ApplyFooter(WordprocessingDocument doc, string input);
    void ApplyHeader(WordprocessingDocument doc, string input);
}