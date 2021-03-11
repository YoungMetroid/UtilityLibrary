using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UtilityLibrary.Loggers
{
	public class TextLogger : Logger
	{
		private static TextLogger textLogger = null;
		private static readonly object padlock = new object();
		private TextLogger()
		{}

		~TextLogger()
		{
		}
		public new static TextLogger getInstance
		{
			get
			{
				lock (padlock)
				{
					if (textLogger == null)
					{
						textLogger = new TextLogger();
					}
					return textLogger;
				}
			}
		}	
		public void clearFile()
		{
			if (!File.Exists(pathLogFile))
			{
				using (StreamWriter sw = File.CreateText(pathLogFile))
				{
					sw.WriteLine("");
				}
			}
		}
	}
}
