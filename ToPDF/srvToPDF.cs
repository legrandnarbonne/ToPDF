using ConnectUNCWithCredentials;
using SimpleLogger;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Threading;
using System.Threading.Tasks;
using ToPDF.Classes;
using static SimpleLogger.SimpleLog;

namespace ToPDF
{
    public partial class srvToPDF : ServiceBase
    {
        Config _config;
        public const string MyServiceName = "ToPDF";
        FileSystemWatcher[] sourceDirectoryWatcher;
        FileSystemWatcher configWatcher;
        List<string> _toDoQueue;
        public string ServicePath;

        LibreOffice _LO;

        int _treated = 0;
        int _detected = 0;

        Task<bool> _workingJob;
        CancellationTokenSource _cancelSource;

        private void evalJobQueue()
        {

            if (_workingJob != null && !_workingJob.IsCompleted) return;

            if (_toDoQueue.Count == 0)
            {
                SimpleLog.Log($"file treated {_treated} detected {_detected}", Severity.Debug);
                return;
            }

            SimpleLog.Log($"file in queue {_toDoQueue.Count}", Severity.Debug);

            _cancelSource = new CancellationTokenSource();
            _cancelSource.CancelAfter(TimeSpan.FromSeconds(120));

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
                }
                return false;
            }, _cancelSource.Token);

            _workingJob.ContinueWith(finishing => {
                SimpleLog.Log($"treatement ended {_toDoQueue.Count} file in queue", Severity.Debug);
                evalJobQueue();
            });
        }

        public srvToPDF()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            ServicePath = AppDomain.CurrentDomain.BaseDirectory;
            _toDoQueue = new List<string>();

            Directory.CreateDirectory(@"C:\Windows\SysWOW64\config\systemprofile\Desktop");

            SimpleLog.SetLogFile(Config.DefaultPath);

            SimpleLog.LogLevel = SimpleLog.Severity.Debug;

            var ps = Process.GetProcessesByName("soffice.exe");
            SimpleLog.Info($"{ps.Length} office process found");


            

            try
            {
                _LO = new LibreOffice();
                _LO.StartOpenOffice();
                SimpleLog.Info("Loading config...");
                _config = Config.Load();

                afterConfigLoaded();

                configWatcher = new FileSystemWatcher(ServicePath, "*.conf");
                configWatcher.NotifyFilter = NotifyFilters.LastWrite;
                configWatcher.Changed += new FileSystemEventHandler(OnConfigChanged);
                configWatcher.EnableRaisingEvents = true;
            }
            catch (Exception e)
            {
                SimpleLog.Error(e.ToString());
            }


        }

        protected override void OnStop()
        {

            disposeWatcher(configWatcher);

            foreach (FileSystemWatcher fw in sourceDirectoryWatcher)
                disposeWatcher(fw);
        }

        private void OnConfigChanged(object sender, FileSystemEventArgs e)
        {
            switch (e.Name.ToLower())
            {
                case "default.conf":
                    SimpleLog.Info("New config détected...");
                    _config = Config.Load();
                    afterConfigLoaded();
                    break;
            }
        }

        void disposeWatcher(FileSystemWatcher fsw)
        {
            if (fsw == null) return;
            fsw.EnableRaisingEvents = false;
            fsw.Dispose();

        }

        void afterConfigLoaded()
        {
            sourceDirectoryWatcher = new FileSystemWatcher[_config.WorkingConfigList.Length];

            for (int i = 0; i < _config.WorkingConfigList.Length; i++)
            {
                var wc = _config.WorkingConfigList[i];
                var dir = new DirectoryInfo(wc.SourceFolder);

                if (dir.Exists)
                {
                    SimpleLog.Info($"Folder found {wc.SourceFolder}");
                    sourceDirectoryWatcher[i] = new FileSystemWatcher(wc.SourceFolder, "*.docx");
                    sourceDirectoryWatcher[i].NotifyFilter = NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.FileName
                                 | NotifyFilters.DirectoryName;
                    sourceDirectoryWatcher[i].Created += new FileSystemEventHandler(OnFileCreated);
                    sourceDirectoryWatcher[i].EnableRaisingEvents = true;
                }
                else
                    SimpleLog.Error($"Folder not found {wc.SourceFolder}");
            }

            SimpleLog.Info($"{_config.WorkingConfigList.Length} File watcher set");
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            var newFile = e.FullPath;

            if (_toDoQueue.Contains(newFile) || Path.GetFileName(newFile).StartsWith("~")) return;//already registred

            _toDoQueue.Add(newFile);
            _detected++;
            SimpleLog.Log($"New file detected :{newFile}. Job queue {_toDoQueue.Count}", Severity.Debug);

            evalJobQueue();
        }

        //bool Analyser(string file)//(object sender, DoWorkEventArgs e)
        //{
        //    //SimpleLog.Info($"{file}");
        //    //var uncs = new List<UNCAccessWithCredentials>();
        //    //file=e.Argument.ToString();
        //    try
        //    {


        //        return convertToPDF(file);


        //    }
        //    catch (Exception ex)
        //    {
        //        SimpleLog.Error(ex.ToString());
        //        return false;
        //    }
        //}

        //private bool convertToPDF(string v)
        //{
        //    WorkingConfig currentWorkingConfig = null;

        //    foreach (var wc in _config.WorkingConfigList)
        //        if (Path.GetDirectoryName(v) == wc.SourceFolder)
        //        {
        //            currentWorkingConfig = wc;
        //            break;
        //        }

        //    if (currentWorkingConfig == null)
        //    {
        //        SimpleLog.Error($"No config find");
        //        openedFile--;
        //        return false;
        //    }

        //    var destPath = currentWorkingConfig.DestinationFolder.TrimEnd('\\') + "\\" + Path.GetFileNameWithoutExtension(v) + ".pdf";

        //    LibreOffice.ConvertToPdf(v, destPath);

        //    SimpleLog.Log("Cleaning", SimpleLog.Severity.Debug);

        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();
        //    // GC needs to be called twice in order to get the Finalizers called  
        //    // - the first time in, it simply makes a list of what is to be  
        //    // finalized, the second time in, it actually is finalizing. Only  
        //    // then will the object do its automatic ReleaseComObject. 
        //    GC.Collect();
        //    GC.WaitForPendingFinalizers();

        //    _treated++;
        //    openedFile--;

        //    evalJobQueu();

        //    return true;

        //}

        private void getPrinterList()
        {
            foreach (string printer in System.Drawing.Printing.PrinterSettings.InstalledPrinters)
            {
                SimpleLog.Info(printer);
            }
        }





    }
}
