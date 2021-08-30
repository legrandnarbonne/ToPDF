using SimpleLogger;
using System;
using System.IO;
using System.Threading.Tasks;
using ToPDF.Classes;

namespace TestToPDF
{
    class Worker
    {
        public Worker(string sourceFile)
        {
            //Analyser(sourceFile);
        }

        //public async Task<bool> Analyser(string file)//(object sender, DoWorkEventArgs e)
        //{
        //    //SimpleLog.Info($"{file}");

        //    //file=e.Argument.ToString();
        //    try
        //    {
        //        //var uncs = new List<UNCAccessWithCredentials>();

        //        if (!FileLock.WaitForFileSize(file) && false)
        //        {
        //            var err = "La taille du fichier " + file + " ne c'est pas stabilisé dans le temps imparti.";
        //            SimpleLog.Error(err);
                    
        //            return false;
        //        }

        //        return convertToPDF(file);


        //    }
        //    catch (Exception ex)
        //    {
        //        SimpleLog.Error(ex.ToString());
        //        return false;
        //    }
        //}


        public static void convertToPDF(Object stateInfo)
        {
            WorkingConfig currentWorkingConfig = null;
            var v = stateInfo.ToString();
            //foreach (var wc in _config.WorkingConfigList)
            //    if (Path.GetDirectoryName(v) == wc.SourceFolder)
            //    {
            //        currentWorkingConfig = wc;
            //        break;
            //    }

            var dest = @"c:\test\pdf";

            //if (currentWorkingConfig == null)
            //{
            //    SimpleLog.Error($"No config find");
            //    return false;
            //}

            var destPath = dest + "\\" + Path.GetFileNameWithoutExtension(v) + ".pdf";

            //LibreOffice2.ConvertToPdf(v, destPath);

            SimpleLog.Log("Cleaning", SimpleLog.Severity.Debug);

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called  
            // - the first time in, it simply makes a list of what is to be  
            // finalized, the second time in, it actually is finalizing. Only  
            // then will the object do its automatic ReleaseComObject. 
            GC.Collect();
            GC.WaitForPendingFinalizers();


            //return true;

        }

    }
}
