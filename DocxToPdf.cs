using System.IO;
using DinkToPdf;
using DinkToPdf.Contracts;



namespace PDFile
{
    public class DocxToPdf : PdfConventer
    {
        private static readonly IConverter Converter
            = new SynchronizedConverter(new PdfTools());

        public override byte[] ToPdf(MemoryStream downloadStream)
        {
            downloadStream.Position = 0;
           var html = DocxToHtml.Docx.ConvertToHtml(downloadStream);
           var doc = new HtmlToPdfDocument
            {
                Objects = {
                    new ObjectSettings {
                        HtmlContent = html ,
                        WebSettings = { DefaultEncoding = "utf-8" }
                    }
                }
            };
            var bytes = Converter.Convert(doc);
            

            return bytes;
        }
    }
}
        