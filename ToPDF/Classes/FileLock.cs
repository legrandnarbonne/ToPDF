using System;
using System.IO;

namespace ToPDF.Classes
{
    static class FileLock
    {
        /// <summary>
        /// Blocks until the file is not locked any more. return true if file is unlock false on time out
        /// Thanks to Eric Z Beard
        /// Source http://stackoverflow.com/questions/50744/wait-until-file-is-unlocked-in-net
        /// </summary>
        /// <param name="fullPath"></param>
        public static bool WaitForFile(string fullPath)
        {
            int numTries = 0;
            while (true)
            {
                ++numTries;
                try
                {
                    // Attempt to open the file exclusively.
                    using (FileStream fs = new FileStream(fullPath,
                        FileMode.Open, FileAccess.ReadWrite,
                        FileShare.None, 100))
                    {
                        fs.ReadByte();
                        break;
                    }
                }
                catch 
                {

                    if (numTries > 50)
                    {
                        return false;
                    }

                    System.Threading.Thread.Sleep(1000);
                }
            }

            return true;
        }

        public static bool WaitForFileSize(string fullPath)
        {
            int numTries = 0;
            var fi = new FileInfo(fullPath);
            var memSize = fi.Length;

            while (true)
            {
                ++numTries;
                if (numTries > 150) return false;

                System.Threading.Thread.Sleep(5000);
                fi = new FileInfo(fullPath);

                if (memSize == fi.Length) return true;
                else memSize = fi.Length;

            }
            
        }
    }
}
