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
    class LibreOffice
    {
        XComponentLoader aLoader;
        int _retry;
        int _maxRetry = 3;

        public bool ConvertToPdf(string inputFile, string outputFile)
        {
            if (ConvertExtensionToFilterType(Path.GetExtension(inputFile)) == null)
                throw new InvalidProgramException("Unknown file type for OpenOffice. File = " + inputFile);

            //Load the sourcefile

            XComponent xComponent = null;
            try
            {
                xComponent = InitDocument(aLoader,
                                        PathConverter(inputFile), "_blank");
                //Wait for loading
                while (xComponent == null)
                    Thread.Sleep(250);

                // save/export the document
                SaveDocument(xComponent, inputFile, PathConverter(outputFile));
            }
            catch (Exception e)
            {
                SimpleLog.Error(e.ToString() + "\r\n restarting libre office.");
                StartOpenOffice(true);

                if (_retry < _maxRetry)
                {
                    SimpleLog.Error($"retry {inputFile}.");
                    _retry++;
                    ConvertToPdf(inputFile, outputFile);
                }
            }
            finally
            {
                _retry = 0;
                if (xComponent != null) xComponent.dispose();
            }

            return true;
        }
        /// <summary>
        /// Start libre office instance if needed
        /// </summary>
        /// <param name="restart">kill previously instance if exist</param>
        public void StartOpenOffice(bool restart = false)
        {
            var ps = Process.GetProcessesByName("soffice.exe");
            if (ps.Length != 0)
                if (restart)
                {
                    foreach (Process pp in ps)
                        pp.Kill();
                    Thread.Sleep(2000);
                }
                else
                    throw new InvalidProgramException("OpenOffice not found.  Is OpenOffice installed?");

            if (ps.Length > 0)
                return;


            var p = new Process
            {
                StartInfo =
                        {
                            Arguments = "--headless --nofirststartwizard --norestore",
                            FileName = "soffice.exe",
                            CreateNoWindow = true
                        }
            };
            var result = p.Start();

            SimpleLog.Log($"Starting office.", SimpleLog.Severity.Info);

            if (result == false)
                SimpleLog.Log($"Starting error result: {result}.", SimpleLog.Severity.Error);

            //Get a ComponentContext
            var xLocalContext =
                Bootstrap.bootstrap();
            //Get MultiServiceFactory
            var xRemoteFactory =
                (XMultiServiceFactory)
                xLocalContext.getServiceManager();
            //Get a CompontLoader
            aLoader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");

        }

        private XComponent InitDocument(XComponentLoader aLoader, string file, string target)
        {
            var openProps = new PropertyValue[1];
            openProps[0] = new PropertyValue { Name = "Hidden", Value = new Any(true) };

            var xComponent = aLoader.loadComponentFromURL(
                file, target, 0,
                openProps);

            return xComponent;
        }

        private void SaveDocument(XComponent xComponent, string sourceFile, string destinationFile)
        {
            var propertyValues = new PropertyValue[2];
            // Setting the flag for overwriting
            propertyValues[1] = new PropertyValue { Name = "Overwrite", Value = new Any(true) };
            //// Setting the filter name
            propertyValues[0] = new PropertyValue
            {
                Name = "FilterName",
                Value = new Any(ConvertExtensionToFilterType(Path.GetExtension(sourceFile)))
            };
            ((XStorable)xComponent).storeToURL(destinationFile, propertyValues);
        }

        private string PathConverter(string file)
        {
            if (string.IsNullOrEmpty(file))
                throw new NullReferenceException("Null or empty path passed to OpenOffice");

            return String.Format("file:///{0}", file.Replace(@"\", "/"));
        }

        public string ConvertExtensionToFilterType(string extension)
        {
            switch (extension)
            {
                case ".doc":
                case ".docx":
                case ".txt":
                case ".rtf":
                case ".html":
                case ".htm":
                case ".xml":
                case ".odt":
                case ".wps":
                case ".wpd":
                    return "writer_pdf_Export";
                case ".xls":
                case ".xlsb":
                case ".xlsx":
                case ".ods":
                    return "calc_pdf_Export";
                case ".ppt":
                case ".pptx":
                case ".odp":
                    return "impress_pdf_Export";

                default:
                    return null;
            }
        }
    }
}
