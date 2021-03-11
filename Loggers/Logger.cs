using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace UtilityLibrary.Loggers
{
    public class Logger
    {
		protected string version = "1.0.0.0";
        protected string pathLogFile = "";// = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private static Logger logger = null;
        private static readonly object padlock = new object();
        
        protected Logger()
        {
		}
        public static Logger getInstance
        {
            get
            {
                lock (padlock)
                {
                    if (logger == null)
                    {
                        logger = new Logger();
                    }
                    return logger;
                }
            }
        }
        public string getLogPath()
        {
            return pathLogFile;
        }
        public void setLogPathandFile(string logPath, string file)
        {
            pathLogFile = logPath + file;
		}
		public void setVersion(string version)
		{
			this.version = version;
		}
		public void startLogger()
		{
			addTextToLogFile("****************************************************************************");
			logDateAndVersion();
		}
		public void endLogger()
		{
			addTextToLogFile("****************************************************************************");
		}
        
        public void logDateAndVersion()
        {
            addTextToLogFile("Date: " + DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss"));
            addTextToLogFile("Version: " + version);
        }

        public void logException(Exception e)
        {
            addTextToLogFile("**************************************");
            logDateAndVersion();
            addTextToLogFile("Type: " + e.GetType().ToString());
            addTextToLogFile("Message: " + e.Message);
            addTextToLogFile("Stack trace: " + e.StackTrace);
            addTextToLogFile("**************************************");
            addTextToLogFile("");
        }
        public void logException(Exception e, string s)
        {
            addTextToLogFile("**************************************");
            logDateAndVersion();
            addTextToLogFile("Type: " + e.GetType().ToString());
            addTextToLogFile("Message: " + e.Message);
            addTextToLogFile("Stack trace: " + e.StackTrace);
            addTextToLogFile("Data:" + s);
            addTextToLogFile("**************************************");
            addTextToLogFile("");
        }

        public void addTextToLogFile(string logMessage)
        {
            if (!File.Exists(pathLogFile))
            {
                using (StreamWriter sw = File.CreateText(pathLogFile))
                {
                    sw.WriteLine(logMessage);
                }
            }
            else
            {
                if (File.GetLastWriteTime(pathLogFile).Date == DateTime.Today.Date)
                {
                    using (StreamWriter sw = File.AppendText(pathLogFile))
                    {
                        sw.WriteLine(logMessage);
                    }
                }
                else
                {
                    try
                    {
                        File.Delete(pathLogFile);
                        using (StreamWriter sw = File.CreateText(pathLogFile))
                        {
                            sw.WriteLine(logMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.Print(ex.Message);
                    }
                }
            }
        }
    }
}
