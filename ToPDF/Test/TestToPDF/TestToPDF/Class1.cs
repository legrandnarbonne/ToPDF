using ConnectUNCWithCredentials;
using SimpleLogger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ToPDF.Classes;
using uno.util;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using static SimpleLogger.SimpleLog;

namespace TestToPDF
{
    class Class1
    {

        Config _config;
        public const string MyServiceName = "ToPDF";
        FileSystemWatcher[] sourceDirectoryWatcher;
        FileSystemWatcher configWatcher;
        List<string> _toDoQueue;
        public string ServicePath;


        int _treated = 0;
        int _detected = 0;

        Task<bool> _workingJob;

        LibreOffice _LO;

        public async void start()
        {
            _toDoQueue = new List<string>();
            _config = Config.Load();

            var d = new DirectoryInfo(@"c:\test");
            _config = Config.Load();

            _LO = new LibreOffice();

            _LO.StartOpenOffice();

            foreach (var fi in d.GetFiles("*.docx"))
            {
                _toDoQueue.Add(fi.FullName);

                evalJobQueue();
            }
        }

        private void evalJobQueue()
        {

            if (_workingJob != null &&!_workingJob.IsCompleted) return;

            if (_toDoQueue.Count == 0)
            {
                SimpleLog.Log($"file treated {_treated} detected {_detected}", Severity.Debug);
                return;
            }

            SimpleLog.Log($"file in queue {_toDoQueue.Count}", Severity.Debug);




            _workingJob = Task.Run(() =>
            {
                var newFile = _toDoQueue[0];
                _treated++;

                var fi = new FileInfo(newFile);

                try
                {
                    // is file avaible
                    if (!fi.Exists)
                        throw new Exception("File not found");

                    if ((fi.CreationTime - DateTime.Now).Seconds > 10 && !FileLock.WaitForFileSize(newFile))
                        throw new Exception("La taille du fichier " + newFile + " ne c'est pas stabilisé dans le temps imparti.");

                    // finding corresponding config
                    WorkingConfig currentWorkingConfig = null;

                    foreach (var wc in _config.WorkingConfigList)
                        if (Path.GetDirectoryName(newFile) == wc.SourceFolder)
                        {
                            currentWorkingConfig = wc;
                            break;
                        }

                    if (currentWorkingConfig == null)
                        throw new Exception($"No config found for this path");

                    var destFile = $"{currentWorkingConfig.DestinationFolder.TrimEnd('\\')}\\{Path.GetFileNameWithoutExtension(newFile)}.pdf";

                    //start convertion

                    var destFi = new FileInfo(destFile);

                    if (!destFi.Exists || currentWorkingConfig.OverWriteIfExist)
                    {
                        _toDoQueue.Remove(newFile);
                        return _LO.ConvertToPdf(newFile, destFile);
                    }

                    if (currentWorkingConfig.DeleteAfterConvert)
                        fi.Delete();

                    _toDoQueue.Remove(newFile);
                    return true;
                }
                catch (Exception e)
                {
                    SimpleLog.Error($"File {newFile} error." + e.ToString());
                    _toDoQueue.Remove(newFile);
                }
                return false;
            });

            _workingJob.ContinueWith(
                finishing => 
                {
                    evalJobQueue(); });
            //_workingJob.Wait();
        }


        //private void evalJobQueu()
        //{
        //    SimpleLog.Log($"Queue control {_toDoQueue.Count}", Severity.Debug);

        //    if (_toDoQueue.Count == 0 || openedFile >= maxOpenedFile) return;

        //    SimpleLog.Log($"waiting {openedFile} opened file, queu {_toDoQueue.Count}, treated {_treated} {openedFile <= maxOpenedFile}", Severity.Info);


        //    var next = _toDoQueue.First();

        //    openedFile++;
        //    _toDoQueue.Remove(next);

        //    Task<bool> taskA = Task.Run(() =>
        //    {
        //        return Analyser(next);
        //    });
        //    taskA.Wait();
        //}

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            if (_toDoQueue.Contains(e.FullPath) || e.Name.StartsWith("~")) return;//already registred
            _toDoQueue.Add(e.FullPath);

            SimpleLog.Log($"New file detected :{e.FullPath}. Job queu {_toDoQueue.Count}", Severity.Debug);

            //evalJobQueu();
        }


    }
}


