using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace UtilityLibrary.ProcessManager
{
	public static class ProcessManager
	{
		private static Logger logger = Logger.getInstance;
		public static void killProcess(string nameOfProcess)
		{
			try
			{
				System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName(nameOfProcess);
				foreach (System.Diagnostics.Process PK in PROC)
				{
					PK.Kill();
				}
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
		}
	}
}
