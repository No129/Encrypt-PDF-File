using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.Word;

namespace Entry
{
    /// <summary>
    /// Interaction logic for MainView.xaml
    /// </summary>
    public partial class MainView : System.Windows.Window
    {
        public MainView()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog objFileDialog = new System.Windows.Forms.OpenFileDialog();

            objFileDialog.Title = "請選取待保護的 PDF 檔案";
            objFileDialog.InitialDirectory = "C:\\";
            if (objFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.FilePathTextBox.Text = objFileDialog.FileName;
                this.SaveAsButton.IsEnabled = true;
            }
            else
            {
                this.SaveAsButton.IsEnabled = false;
            }
        }

        private void WordFileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog objFileDialog = new System.Windows.Forms.OpenFileDialog();

            objFileDialog.Title = "請選取轉換 Word 檔案";
            objFileDialog.InitialDirectory = "C:\\";
            if (objFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.WordFilePathTextBox.Text = objFileDialog.FileName;
                this.SaveAsButton.IsEnabled = true;
            }
            else
            {
                this.SaveAsButton.IsEnabled = false;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.SetPDFFile(this.FilePathTextBox.Text);
            MessageBox.Show("完成指定設定。");
        }

        private void SaveAsPDFButton_Click(object sender, RoutedEventArgs e)
        {   
            Microsoft.Office.Interop.Word.Application objApp = null;
            Microsoft.Office.Interop.Word.Document objDoc = null;
            string sWordPath = this.WordFilePathTextBox.Text;
            string sPDFPath = this.GetPDFFilePath(sWordPath);

            objApp = new Microsoft.Office.Interop.Word.Application();
            objDoc = objApp.Documents.Open(sWordPath);
            objDoc.ExportAsFixedFormat(sPDFPath,
                Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                false ,
                Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument,
                IncludeDocProps:true ,
                BitmapMissingFonts:true,
                Item:Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent);

            objDoc.Close();
            objApp.Quit(SaveChanges: false);
            this.SetPDFFile(sPDFPath);
            MessageBox.Show("完成 Word 檔案輸出。");
            System.Diagnostics.Process.Start(sPDFPath);            
        }

        private void IsNeedPassWordForOpenFileCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            switch (this.IsNeedPassWordForOpenFileCheckBox.IsChecked)
            {
                case true:
                    this.PasswordTextBox.IsEnabled = true;
                    break;

                case false:
                    this.PasswordTextBox.Text = string.Empty;
                    this.PasswordTextBox.IsEnabled = false;
                    break;
            }
        }

        private string GetPDFFilePath(string pi_sWordFilePath)
        {
            string sReturn = string.Empty;
            string sDirectory = System.IO.Path.GetDirectoryName(pi_sWordFilePath);
            string sFileName = System.IO.Path.GetFileNameWithoutExtension(pi_sWordFilePath);

            sReturn = string.Format("{0}\\{1}.pdf", sDirectory, sFileName);
            if(System.IO.File.Exists(sReturn))
            {
                System.IO.File.Delete(sReturn);
            }

            return sReturn;
        }

        private void SetPDFFile(string sPDFFilePath)
        {
            TargetFile objTarget = new TargetFile(sPDFFilePath);
            string sPW = "TOHU";
            byte[] objPW = Encoding.ASCII.GetBytes(sPW);

            using (PdfReader objReader = new PdfReader(objTarget.SourcePath, objPW))
            {
                using (var objOutputFileStream = new FileStream(objTarget.OutputPath, FileMode.Create))
                {
                    string sPassword = this.IsNeedPassWordForOpenFileCheckBox.IsChecked == true ? this.PasswordTextBox.Text : null;
                    int nPermission = 0;

                    if (this.AllowCopyCheckBox.IsChecked == true)
                    {
                        nPermission = nPermission | PdfWriter.AllowCopy;
                    }

                    if (this.AllowPrintingCheckBox.IsChecked == true)
                    {
                        nPermission = nPermission | PdfWriter.AllowPrinting;
                    }
                    PdfEncryptor.Encrypt(objReader, objOutputFileStream, true, sPassword, sPW, nPermission);
                }
            }
            objTarget.CleanTempFile();
        }
    }
}
