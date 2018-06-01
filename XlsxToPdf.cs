using System.IO;
using DinkToPdf;
using DinkToPdf.Contracts;
using NPOI.SS.Converter;
using NPOI.SS.UserModel;

namespace PDFile
{
    public class XlsxToPdf:PdfConventer
    {
        private static readonly IConverter Converter
            = new SynchronizedConverter(new PdfTools());
        
        public override byte[] ToPdf(MemoryStream downloadStream)
                      {
                          downloadStream.Position = 0;
                          var excelToHtmlConverter = new ExcelToHtmlConverter
                          {
                              OutputColumnHeaders = false,
                              OutputHiddenColumns = false,
                              OutputHiddenRows = false,
                              OutputLeadingSpacesAsNonBreaking = false,
                              OutputRowNumbers = false,
                              UseDivsToSpan = false
                          };
                          excelToHtmlConverter.ProcessWorkbook(WorkbookFactory.Create(downloadStream)); 
                          
                          var doc = new HtmlToPdfDocument
                          {   GlobalSettings =
                              {
                                  Orientation = Orientation.Landscape
                              },
                              Objects = {
                                  new ObjectSettings {
                                      HtmlContent = excelToHtmlConverter.Document.OuterXml,
                                      WebSettings = { DefaultEncoding = "utf-8" }
                                  }
                              }
                          };
          
            var bytes = Converter.Convert(doc);
            return bytes;
        }
    }
}