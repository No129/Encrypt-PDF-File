using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Text;
using System.Windows;
using Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace Entry
{
    /// <summary>
    /// Interaction logic for MainView.xaml
    /// </summary>
    public partial class MainView : System.Windows.Window
    {
        #region -- 建構/解構 ( Constructors/Destructor ) --

        public MainView()
        {
            InitializeComponent();
        }

        #endregion

        #region -- 事件處理 ( Event Handlers ) --

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

        private void MergeWordFileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog objFileDialog = new System.Windows.Forms.OpenFileDialog();

            objFileDialog.Title = "請選取合併 Word 檔案";
            objFileDialog.InitialDirectory = "C:\\";
            if (objFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.MergeWordFilePathTextBox.Text = objFileDialog.FileName;
                this.MergeAsPDFButton.IsEnabled = true;
            }
            else
            {
                this.MergeAsPDFButton.IsEnabled = false;
            }
        }

        private void MergePDFFileSelectButton_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog objFileDialog = new System.Windows.Forms.OpenFileDialog();

            objFileDialog.Title = "請選取合併 Word 檔案";
            objFileDialog.InitialDirectory = "C:\\";
            if (objFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.MergePDFFilePathTextBox.Text = objFileDialog.FileName;
                this.MergeAsPDFButton.IsEnabled = true;
            }
            else
            {
                this.MergeAsPDFButton.IsEnabled = false;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.SetPDFFile(this.FilePathTextBox.Text);
            MessageBox.Show("完成指定設定。");
        }

        private void SaveAsPDFButton_Click(object sender, RoutedEventArgs e)
        {
            string sWordPath = this.WordFilePathTextBox.Text;
            string sPDFPath = new PDFHelper().FromWord(sWordPath);

            if (string.IsNullOrEmpty(sPDFPath) != true)
            {
                this.SetPDFFile(sPDFPath);  //設定檔案保護。
                MessageBox.Show("完成 Word 檔案輸出。");
                System.Diagnostics.Process.Start(sPDFPath); //打開 PDF 文件。
            }
        }

        private void MergeAsPDFButton_Click(object sender, RoutedEventArgs e)
        {
            List<string> objTarget = new List<string>();
            string sWordPath = this.MergeWordFilePathTextBox.Text;
            string sPDFPath = new PDFHelper().FromWord(sWordPath);

            if (string.IsNullOrEmpty(sPDFPath) != true)
            {
                objTarget.Add(sPDFPath);
                if (string.IsNullOrEmpty(this.MergePDFFilePathTextBox.Text) != true)
                {
                    objTarget.Add(this.MergePDFFilePathTextBox.Text);
                    string sPDFFile = new PDFHelper().MergePDF(objTarget);
                    MessageBox.Show("完成檔案整合。");
                    System.Diagnostics.Process.Start(sPDFFile); //打開 PDF 文件。
                }
            }
        }

        #endregion

        #region -- 私有函式 ( Private Method) --

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

        #endregion

    }
}
