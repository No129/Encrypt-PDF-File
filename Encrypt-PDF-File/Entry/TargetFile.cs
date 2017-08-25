using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entry
{
    public class TargetFile
    {
        string l_sSourcePath = string.Empty;
        string l_sOutputPath = string.Empty;

        public TargetFile(string pi_sPath)
        {
            this.l_sOutputPath = pi_sPath;
            this.l_sSourcePath = this.CopyTemplationFile(this.l_sOutputPath);
        }
       
        public void CleanTempFile()
        {
            new System.IO.FileInfo(this.l_sSourcePath).Delete();
        }

        public string SourcePath { get { return this.l_sSourcePath; } }
        public string OutputPath { get { return this.l_sOutputPath; } }

        private string CopyTemplationFile(string pi_sOutputPath)
        {
            string sReturn = string.Empty;
            string sDirectory = System.IO.Path.GetDirectoryName(pi_sOutputPath);            
         
            sReturn = string.Format("{0}\\{1}.pdf", sDirectory, Guid.NewGuid().ToString());
            System.IO.File.Copy(pi_sOutputPath, sReturn);

            return sReturn;
        }
    }
}
