using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Entry
{
    public class PDFHelper
    {  
        #region -- 方法 ( Public Method ) --

        /// <summary>
        /// 從 Word 轉出 PDF 檔案。
        /// </summary>
        /// <param name="pi_sWordPath">待轉 Word 檔案路徑。</param>
        /// <returns>PDF 檔案路徑。</returns>
        public string FromWord(string pi_sWordPath)
        {
            string sReturn = this.GetPDFFilePath(pi_sWordPath);

            if (string.IsNullOrEmpty(sReturn) != true)
            {
                Microsoft.Office.Interop.Word.Application objApp = null;
                Microsoft.Office.Interop.Word.Document objDoc = null;

                objApp = new Microsoft.Office.Interop.Word.Application();
                objDoc = objApp.Documents.Open(pi_sWordPath);

                objDoc.ExportAsFixedFormat(sReturn,
                    Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                    false,
                    Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument,
                    IncludeDocProps: true,
                    BitmapMissingFonts: true,
                    Item: Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent);

                objDoc.Close();
                objApp.Quit(SaveChanges: false);
            }

            return sReturn;
        }

        public string MergePDF(List<string> pi_objTarget)
        {
            string sReturn = string.Format("C://test-{0}.pdf", System.DateTime.Now.ToString("yyMMdd-hhmmss"));

            // step 1: creation of a document-object
            Document document = new Document();
            // step 2: we create a writer that listens to the document
            PdfCopy writer = new PdfCopy(document, new FileStream(sReturn, FileMode.Create));

            try
            {
                if (writer != null)
                {
                    // step 3: we open the document
                    document.Open();

                    foreach (string fileName in pi_objTarget)
                    {
                        // we create a reader for a certain document
                        PdfReader reader = new PdfReader(fileName);
                        reader.ConsolidateNamedDestinations();

                        // step 4: we add content
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {                          
                            PdfImportedPage page = writer.GetImportedPage(reader, i);
                            writer.AddPage(page);                            
                        } 
                        reader.Close();
                    }
                    // step 5: we close the document and writer
                    writer.Close();
                    document.Close();
                }
            }
            finally
            {
                writer = null;
                document = null;
            }
           
            return sReturn;
        }

        #endregion        

        #region -- 私有函式 ( Private Method) --

        /// <summary>
        /// 取得對應 Word 的 PDF 檔名。
        /// </summary>
        /// <param name="pi_sWordFilePath">Word 檔名。</param>
        /// <returns>對應 Word 的 PDF 檔名。</returns>
        private string GetPDFFilePath(string pi_sWordFilePath)
        {
            string sReturn = string.Empty;
            string sDirectory = System.IO.Path.GetDirectoryName(pi_sWordFilePath);
            string sFileName = System.IO.Path.GetFileNameWithoutExtension(pi_sWordFilePath);

            sReturn = string.Format("{0}\\{1}.pdf", sDirectory, sFileName);
            if (System.IO.File.Exists(sReturn))
            {
                try
                {
                    System.IO.File.Delete(sReturn);
                }
                catch
                {
                    MessageBox.Show("請先關閉先前開啟的 PDF 檔案。");
                    sReturn = string.Empty;
                }
            }

            return sReturn;
        }

        #endregion
           
    }
}
