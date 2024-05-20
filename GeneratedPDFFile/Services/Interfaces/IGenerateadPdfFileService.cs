namespace GeneratedPDFFile.Services.Interfaces;

public interface IGenerateadPdfFileService
{
     ValueTask InsertHtmlToWord(string htmlContent);
     ValueTask EskiHolat(string htmlstring);
     void ConvertHtmlToDocx(string html);
}