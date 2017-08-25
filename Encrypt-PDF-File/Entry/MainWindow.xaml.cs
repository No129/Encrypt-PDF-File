using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace Entry
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.SaveAsButton.IsEnabled = false;
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TargetFile objTarget = new TargetFile(this.FilePathTextBox.Text);
            string sPW = "TOHU";
            byte[] objPW = Encoding.ASCII.GetBytes(sPW);

            using (PdfReader objReader = new PdfReader(objTarget.SourcePath, objPW))
            {
                using (var objOutputFileStream = new FileStream(objTarget.OutputPath, FileMode.Create))
                {
                    string sPassword = this.IsNeedPassWordForOpenFileCheckBox.IsChecked == true ? this.PasswordTextBox.Text : null;
                    int nPermission = 0;

                    if(this.AllowCopyCheckBox.IsChecked == true)
                    {
                        nPermission = nPermission | PdfWriter.AllowCopy;
                    }

                    if(this.AllowPrintingCheckBox.IsChecked == true)
                    {
                        nPermission = nPermission | PdfWriter.AllowPrinting;
                    }
                    PdfEncryptor.Encrypt(objReader, objOutputFileStream, true, sPassword, sPW, nPermission);                   
                }
            }
            objTarget.CleanTempFile();

            MessageBox.Show("完成指定設定。");
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
    }
}
