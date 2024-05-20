using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GeneratedPDFFile.Entities;
using GeneratedPDFFile.Services.Interfaces;
using Xceed.Words.NET;
using Document = Xceed.Document.NET.Document;
using Paragraph = Xceed.Document.NET.Paragraph;

namespace GeneratedPDFFile.Services;
public class GenerateadPdfFileService : IGenerateadPdfFileService
{
    // public async ValueTask InsertHtmlToWord(string htmlContent)
    // {
    //     
    //     string filePath = @"C:\File\TpiFile.docx";
    //     DocX document;
    //     if (File.Exists(filePath))
    //     {
    //         document = DocX.Load(filePath);
    //     }
    //     else
    //     {
    //         document = DocX.Create(filePath);
    //     }
    //     
    //     Paragraph startMarker = document.Paragraphs.FirstOrDefault(p => p.Text.Contains("[START]"));
    //     Paragraph endMarker = document.Paragraphs.FirstOrDefault(p => p.Text.Contains("[END]"));
    //
    //
    //     if (startMarker != null && endMarker != null)
    //     {
    //         int startIndex = document.Paragraphs.IndexOf(startMarker);
    //         int endIndex = document.Paragraphs.IndexOf(endMarker);
    //
    //         for (int i = startIndex; i < endIndex; i++)
    //         {
    //             document.InsertParagraph(htmlContent);
    //         }
    //
    //     }
    //     
    //     int numberOfLinesToAdd = 0;
    //     if (document.Paragraphs.Count > 0)
    //     {
    //         numberOfLinesToAdd = document.Paragraphs[^1].Text
    //             .Split(new[] { Environment.NewLine }, StringSplitOptions.None).Length;
    //     }
    //
    //     for (int i = 0; i < numberOfLinesToAdd; i++)
    //     {
    //         document.InsertParagraph(Environment.NewLine);
    //     }
    //     
    //     document.InsertParagraph(htmlContent);
    //
    //     document.Save();
    // }


    public async ValueTask InsertHtmlToWord(string htmlContent)
    {
        throw new NotImplementedException();
    }

    public async ValueTask EskiHolat(string htmlstring)
    {
        string FilePath = @"C:\File\File21.docx";


        FileStream fileStream = new FileStream(FilePath, FileMode.Append);

        for (int i = 0; i < 18; i++)
        {
            htmlstring += Environment.NewLine;
        }
            
        byte[] bytedata = Encoding.Default.GetBytes(htmlstring);
            
        fileStream.Write( bytedata, 0, bytedata.Length);
            
        fileStream.Close();
    }

    public void ConvertHtmlToDocx(string html)
    {
        throw new NotImplementedException();
    }


    // public void ConvertHtmlToDocx(string html)
    // {
    //     string filePath = @"C:\File\File21.docx";
    //     using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
    //     {
    //         MainDocumentPart mainPart = wordDoc.MainDocumentPart;
    //
    //         
    //         string bookmarkStartId = "start";
    //         string bookmarkEndId = "end";
    //
    //         
    //         BookmarkStart bookmarkStart = mainPart.Document.Body.Descendants<BookmarkStart>().FirstOrDefault(b => b.Name == bookmarkStartId);
    //         BookmarkEnd bookmarkEnd = mainPart.Document.Body.Descendants<BookmarkEnd>().FirstOrDefault(b => b.Id == bookmarkStart.Id);
    //
    //         if (bookmarkStart != null && bookmarkEnd != null)
    //         {
    //             
    //             var currentElement = bookmarkStart.NextSibling();
    //             while (currentElement != null && currentElement != bookmarkEnd)
    //             {
    //                 var nextElement = currentElement.NextSibling();
    //                 currentElement.Remove();
    //                 currentElement = nextElement;
    //             }
    //
    //             
    //             HtmlConverter converter = new HtmlConverter(mainPart);
    //            var paragraphs = converter.Parse(html);
    //
    //             foreach (OpenXmlElement paragraph in paragraphs)
    //             {
    //                 bookmarkStart.InsertAfterSelf(paragraph);
    //             }
    //
    //             mainPart.Document.Save();
    //         }
    //         else
    //         {
    //             throw new InvalidOperationException("Bookmark start or end not found in the document.");
    //         }
    //     }
    // }
    
    private static Dictionary<string, BookmarkEnd> FindBookmarks(OpenXmlElement documentPart, Dictionary<string, BookmarkEnd> results = null, Dictionary<string, string> unmatched = null )
    {
        results = results ?? new Dictionary<string, BookmarkEnd>();
        unmatched = unmatched ?? new Dictionary<string,string>();

        foreach (var child in documentPart.Elements())
        {
            if (child is BookmarkStart)
            {
                var bStart = child as BookmarkStart;
                unmatched.Add(bStart.Id, bStart.Name);
            }

            if (child is BookmarkEnd)
            {
                var bEnd = child as BookmarkEnd;
                foreach (var orphanName in unmatched)
                {
                    if (bEnd.Id == orphanName.Key)
                        results.Add(orphanName.Value, bEnd);
                }
            }

            FindBookmarks(child, results, unmatched);
        }

        return results;
    }

}