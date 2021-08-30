using System.ServiceProcess;

namespace ToPDF
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new srvToPDF()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
