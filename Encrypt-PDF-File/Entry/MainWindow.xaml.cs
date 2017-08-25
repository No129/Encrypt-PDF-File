using iTextSharp.text.pdf;
using System.IO;
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
            if (objFileDialog.ShowDialog()==System.Windows.Forms.DialogResult.OK)
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
            string sSourceFilePath = this.FilePathTextBox.Text;
            string sOutputFilePath = "C:\\test.pdf";

            using (PdfReader objReader = new PdfReader(sSourceFilePath))
            {
                using (var os = new FileStream(sOutputFilePath, FileMode.Create))
                {
                    PdfEncryptor.Encrypt(objReader, os, true, "", "", PdfWriter.ALLOW_SCREENREADERS);
                }

            }
        }
    }
}
