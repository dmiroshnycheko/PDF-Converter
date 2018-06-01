using System.IO;

namespace PDFile
{
     public abstract  class PdfConventer
    {
        public abstract byte[] ToPdf(MemoryStream ds);
    }
}