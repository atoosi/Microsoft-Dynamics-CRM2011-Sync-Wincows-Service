using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Sync
{
    public static class LogWriter
    {
        public static void WriteErrorLog(Exception ex)
        {
            StreamWriter sw = null;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            try 
            {


                string logFilePath = AppDomain.CurrentDomain.BaseDirectory;
                logFilePath = logFilePath + "Log-" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
                logFileInfo = new FileInfo(logFilePath);
                logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
                if (!logDirInfo.Exists) logDirInfo.Create();
                if (!logFileInfo.Exists)
                {
                    fileStream = logFileInfo.Create();
                }
                else
                {
                    fileStream = new FileStream(logFilePath, FileMode.Append);
                }
                sw = new StreamWriter(fileStream);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + "; " + ex.Message.ToString().Trim());
                sw.Close();

                //sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt",true);
                //sw.WriteLine(DateTime.Now.ToString()+": " + ex.Source.ToString().Trim()+ "; "+ex.Message.ToString().Trim());
                //sw.Flush();
                //sw.Close();
            }
            catch 
            {
            }
        }

        public static void WriteErrorLog(string message)
        {
            StreamWriter sw = null;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;

            try
            {

                string logFilePath = AppDomain.CurrentDomain.BaseDirectory +"\\Logs\\";
                logFilePath = logFilePath + "Log-" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
                logFileInfo = new FileInfo(logFilePath);
                logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
                if (!logDirInfo.Exists) logDirInfo.Create();
                if (!logFileInfo.Exists)
                {
                    fileStream = logFileInfo.Create();
                }
                else
                {
                    fileStream = new FileStream(logFilePath, FileMode.Append);
                }
                sw = new StreamWriter(fileStream);
                sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                sw.Close();


                //sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\LogFile.txt", true);
                //sw.WriteLine(DateTime.Now.ToString() + ": " + message);
                //sw.Flush();
                //sw.Close();
            }
            catch
            {
            }
        }
    }
}
