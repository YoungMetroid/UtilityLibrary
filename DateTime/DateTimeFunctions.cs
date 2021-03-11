using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace UtilityLibrary
{
	public class DateTimeFunctions
	{
		public static Logger logger = Logger.getInstance;
		static public DateTime getCentralStandardTime()
		{
			DateTime timeUtc = DateTime.UtcNow;
			try
			{
				TimeZoneInfo cstZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
				DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, cstZone);
				return cstTime;
			}
			catch (TimeZoneNotFoundException timeZoneNotFoundEx)
			{
				Console.WriteLine("The registry does not define the Central Standard Time zone.");
				logger.logException(timeZoneNotFoundEx);
				return DateTime.UtcNow;
			}
			catch (InvalidTimeZoneException invalidTimeZoneEx)
			{
				Console.WriteLine("Registry data on the Central Standard Time zone has been corrupted.");
				logger.logException(invalidTimeZoneEx);
				return DateTime.UtcNow;
			}
		}
		static public DateTime convertToCentralStandardTime(DateTime date)
		{
			try
			{
				TimeZoneInfo cstZone = TimeZoneInfo.FindSystemTimeZoneById("Central Standard Time");
				DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(date, cstZone);
				return cstTime;
			}
			catch (TimeZoneNotFoundException timeZoneNotFoundEx)
			{
				Console.WriteLine("The registry does not define the Central Standard Time zone.");
				logger.logException(timeZoneNotFoundEx);
				return DateTime.UtcNow;
			}
			catch (InvalidTimeZoneException invalidTimeZoneEx)
			{
				Console.WriteLine("Registry data on the Central Standard Time zone has been corrupted.");
				logger.logException(invalidTimeZoneEx);
				return DateTime.UtcNow;
			}
		}
	}
}
