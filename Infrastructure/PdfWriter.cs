using SelectPdf;
using System.Threading.Tasks;

namespace TogglReport.ConsoleApp.Infrastructure {
    public class PdfWriter {
        public async Task WritePdfFile() {
            string text = @"<html>
                         <body>
                          Hello World from selectpdf.com.
                         </body>
                        </html>
                        ";

            HtmlToPdf converter = new HtmlToPdf();
            // create a new pdf document converting an url
            PdfDocument doc = converter.ConvertHtmlString(text);

            // save pdf document
            doc.Save("Sample.pdf");

            // close pdf document
            doc.Close();
        }
    }
}
