using System;
using System.IO;
using DinkToPdf;


namespace PDFile
{
    public class DocxToPdf : PdfConventer
    {
        public override byte[] ToPdf(MemoryStream downloadStream)
        {
            downloadStream.Position = 0;
            var html = DocxToHtml.Docx.ConvertToHtml(downloadStream);
            Console.WriteLine("Coverted MemoryStream docx в html");
            var doc = new HtmlToPdfDocument
            {
                Objects =
                {
                    new ObjectSettings
                    {
                        HtmlContent = html,
                        WebSettings = {DefaultEncoding = "utf-8"}
                    }
                }
            };
            Console.WriteLine("Clear html");
            var bytes = Converter.Convert(doc);
            Console.WriteLine("Convert html to pdf");
            return bytes;
        }
    }
}