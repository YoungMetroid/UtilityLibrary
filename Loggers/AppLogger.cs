using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UtilityLibrary.AppLogger
{
    public class AppLogger
    {
        private static AppLogger appLogger = null;
        private static readonly object padlock = new object();
        private ToolStripStatusLabel statusLabel;

        private AppLogger()
        {

        }

        public static AppLogger getInstance
        {
            get
            {
                lock(padlock)
                {
                    if(appLogger == null)
                    {
                        appLogger = new AppLogger();
                    }
                    return appLogger;
                }
            }
        }

        public void setStatusLabel(ToolStripStatusLabel statusLabel)
        {
            this.statusLabel = statusLabel;
        }
        
        public void printException(Exception ex)
        {
            statusLabel.Text = ex.Message;
        }
        public void printException(Exception ex, string customMessage)
        {
            statusLabel.Text = customMessage + ": " + ex.Message;
        }
        public void printCurrentAction(string message)
        {
            statusLabel.Text = message;
        }
    }
}
