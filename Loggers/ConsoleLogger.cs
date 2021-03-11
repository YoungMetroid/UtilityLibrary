using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilityLibrary.Loggers
{
	public class ConsoleLogger
	{
		private static ConsoleLogger consoleLogger = null;
		private static readonly object padlock = new object();

		protected ConsoleLogger()
		{

		}
		public static ConsoleLogger getInstance
		{
			get{
				lock(padlock)
				{
					if(consoleLogger == null)
					{
						consoleLogger = new ConsoleLogger();
					}
					return consoleLogger;
				} 
			}
		}
		public void logError(string errorMessage)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine(errorMessage);
		}
		public void logMessage(string message)
		{
			Console.ForegroundColor = ConsoleColor.White;
			Console.WriteLine(message);
		}
		public void logMessageHighLight(string message)
		{
			Console.ForegroundColor = ConsoleColor.Cyan;
			Console.WriteLine(message);
		}
		public void logImportantMessage(string message)
		{
			Console.ForegroundColor = ConsoleColor.Yellow;
			Console.WriteLine(message);
		}
	}
}
