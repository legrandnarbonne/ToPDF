using SimpleLogger;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using uno;
using uno.util;
using unoidl.com.sun.star.beans;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;

namespace ToPDF.Classes
{
    class LibreOffice2
    {
        public static void ConvertToPdf(object o)//string srcFile, string destFile)
        {
            //var ps = Process.GetProcessesByName("soffice.exe");
            //--infilter="Microsoft Word 2007/2010/2013 XML"
            // -env:UserInstallation=file:///tmp/LibreOffice_Conversion_${USER} 
            var srcFile = o.ToString();
            using (var p = new Process
            {
                StartInfo =
                        {
                            //WorkingDirectory=@"C:\Users\dapojero\Downloads\LibreOfficePortable\App\libreoffice\program\",
                            Arguments = $"--headless --nofirststartwizard --norestore --convert-to pdf:writer_pdf_Export --outdir \"c:\\test\\pdf\" " + srcFile,
                            FileName = @"C:\Users\dapojero\Downloads\LibreOfficePortable\App\libreoffice\program\"+"soffice.exe",
                            CreateNoWindow = false,
                            RedirectStandardOutput=true,
                            RedirectStandardError=true,
                            UseShellExecute=false
                        }
            })
            {
                
                
                var result = p.Start();

                string stdout = p.StandardOutput.ReadToEnd();
                string stderrx = p.StandardError.ReadToEnd();

                if (result == false)
                    SimpleLog.Log($"Starting error result: {result}.", SimpleLog.Severity.Error);

                p.WaitForExit();

                if (p.ExitCode != 0)
                    Console.WriteLine(p.ExitCode + " " + srcFile+" "+stdout+" "+stderrx );
                //return p.ExitCode == 0;

            }
        }

    }
}
