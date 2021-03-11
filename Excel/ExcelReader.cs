using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;
using System.Text.RegularExpressions;
using System.Linq;
using ExcelR = Microsoft.Office.Interop.Excel;

namespace UtilityLibrary.Excel
{
	public static class ExcelReader
	{
		private static Application app;
		private static Workbook workbook;
		private static Logger logger = Logger.getInstance;
		private static TextLogger textLogger = TextLogger.getInstance;
		public static Worksheet worksheet;
		public const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

		public static Dictionary<string, int> CreateDictionaryHeader(Worksheet sheet, int headerRow)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			Range headerCell = sheet.Range["A" + headerRow];

			int value;
			string key = "";
			int col = 1;
			while (!string.IsNullOrEmpty(sheet.Cells[headerRow, col].Value2))
			{
				bool KeyExists = dictionary.TryGetValue(sheet.Cells[headerRow, col].Value2, out value);
				if (!KeyExists)
				{
					key = sheet.Cells[headerRow, col].Value2;
					dictionary.Add(key.Trim(), col);
				}
				col++;
			}
			return dictionary;
		}
		public static Dictionary<string, int> CreateDictionaryDoubleHeader(Worksheet sheet, int headerRow1, int headerRow2,bool startIndexZero)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			int value;
			string key = "";
			int column = 1;

			int columnsUsed = sheet.UsedRange.Columns.Count;
			int indexStart = 0;
			if (startIndexZero) indexStart = 1;
			while (columnsUsed >= column)
			{
				bool KeyExists = dictionary.TryGetValue(sheet.Cells[headerRow1, column].Value2 + sheet.Cells[headerRow2, column].Value2, out value);
				if (!KeyExists)
				{
					key = sheet.Cells[headerRow1, column].Value2 + sheet.Cells[headerRow2, column].Value2;
					dictionary.Add(key.Trim(), column - indexStart);
				}
				column++;
			}
			return dictionary;
		}
		public static Dictionary<string, int> CreateHeaderFromArrayWithIndex(object[,] arrayItem, int rowIndex)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			int value;
			string key;
			for (int col = 0; col <= arrayItem.GetUpperBound(1); col++)
			{
				if (arrayItem[rowIndex, col] != null && !String.IsNullOrWhiteSpace(arrayItem[rowIndex, col].ToString()))
				{
					bool KeyExists = dictionary.TryGetValue(arrayItem[rowIndex, col].ToString(), out value);
					if (!KeyExists)
					{
						key = arrayItem[rowIndex, col].ToString().Trim();
						dictionary.Add(key, col);
					}
				}
			}
			return dictionary;
		}
		public static Dictionary<string, int> CreateHeaderFromList(List<string[]> list)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			int value;
			string key;
			for (int counter = 0; counter < list.Count; counter++)
			{
				if (list[counter] != null)
				{
					bool KeyExists = dictionary.TryGetValue(list[counter][0], out value);
					if (!KeyExists)
					{
						key = list[counter][0].Trim();
						dictionary.Add(key, counter);
					}
				}
			}
			return dictionary;
		}
		public static Dictionary<string, int> CreateHeaderFromStringArray(string[] list)
		{
			Dictionary<string, int> dictionary = new Dictionary<string, int>();
			int value;
			string key;
			for (int counter = 0; counter < list.Length; counter++)
			{
				if (list[counter] != null)
				{
					bool KeyExists = dictionary.TryGetValue(list[counter], out value);
					if (!KeyExists)
					{
						key = list[counter].Trim();
						dictionary.Add(key, counter);
					}
				}
			}
			return dictionary;
		}
		public static List<KeyValuePair<object, object>> CreateHeaderListWithIndex(object[,] arrayItem, int rowIndex)
		{
			List<KeyValuePair<object, object>> headerList = new List<KeyValuePair<object, object>>();

			for (int column = 0; column <= arrayItem.GetUpperBound(1); column++)
			{
				if (arrayItem[rowIndex, column] != null && !String.IsNullOrWhiteSpace(arrayItem[rowIndex, column].ToString()))
				{
					KeyValuePair<object, object> item = new KeyValuePair<object, object>(arrayItem[rowIndex, column].ToString().Trim(), column);
					headerList.Add(item);
				}
			}
			return headerList;
		}
		public static void updatePivotTable(string filePath)
		{
			Application app = new Application();
			Workbook workBook = app.Workbooks.Open(filePath, ReadOnly: false);
		}
		public static void convertToXlsx(string filePath, string savePath)
		{
			try
			{
				app = new Application();
				app.DisplayAlerts = false;
				workbook = app.Workbooks.Open(filePath, ReadOnly: false);
				workbook.SaveAs(savePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, null,
				null, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
				Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
				null, null, null);
				releaseMemory();
				killExcel();
			}
			catch(Exception ex)
			{
				killExcel();
				logger.logException(ex);
			}
		}
		public static void convertToXlsb(string filePath, string savePath)
		{
			try
			{
				textLogger.addTextToLogFile("Started Converting File: " + filePath + " to " + "extension .xlsb");
				app = new Application();
				app.DisplayAlerts = false;
				workbook = app.Workbooks.Open(filePath, ReadOnly: false);
				workbook.SaveAs(savePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12, Missing.Value,
				Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
				Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
				Missing.Value, Missing.Value, Missing.Value);
				releaseMemory();
				killExcel();
				logger.addTextToLogFile("Finished Converting File: " + filePath + " to " + "extension .xlsb");
			}
			catch (Exception ex)
			{
				logger.logException(ex);
				killExcel();
			}


		}
		public static void convertToXlsm(string filePath, string savePath)
		{
			try
			{
				logger.addTextToLogFile("Started Converting File: " + filePath + " to " + "extension .xlsm");
				app = new Application();
				app.DisplayAlerts = false;
				workbook = app.Workbooks.Open(filePath, ReadOnly: false);
				workbook.SaveAs(savePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Missing.Value,
				Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
				Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true,
				Missing.Value, Missing.Value, Missing.Value);
				releaseMemory();
				killExcel();
				logger.addTextToLogFile("Finished Converting File: " + filePath + " to " + "extension .xlsm");
			}
			catch(Exception ex)
			{
				logger.logException(ex);
				killExcel();
			}

			
		}
		public static string  convertColumnNumberToLetter(int value)
		{
			StringBuilder builder = new StringBuilder();
			do
			{
				if (builder.Length > 0)
					value--;
				builder.Insert(0, alphabet[value % alphabet.Length]);
				value /= alphabet.Length;
			} while (value > 0);

			return builder.ToString();
		}
		public static void createExcelFile()
		{
			try
			{
				app = new Application();
				workbook = app.Workbooks.Add(System.Reflection.Missing.Value);
				worksheet = workbook.Worksheets.get_Item(1);
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
		}
		public static void killExcel()
		{
			try
			{
				System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
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
		public static void releaseMemory()
		{
			try
			{
				workbook.Close();
				app.Quit();
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}

			if (worksheet != null)
				while (Marshal.ReleaseComObject(worksheet) != 0) { }
			
			if (workbook != null)
				while (Marshal.ReleaseComObject(workbook) != 0) { }
			if (app != null)
				while (Marshal.ReleaseComObject(app) != 0) { }

			GC.Collect();
			GC.WaitForPendingFinalizers();
			worksheet = null;
			workbook = null;
			app = null;
		}
		public static void setFormat(int rowCount, int columnCount)
		{
			Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, columnCount]];
			worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, range, Type.Missing, XlYesNoGuess.xlYes, Type.Missing).Name = "query";
			worksheet.ListObjects["query"].TableStyle = "TableStyleMedium2";
			Marshal.ReleaseComObject(range);
		}
		public static void saveFileAsXlsb(string fileName)
		{
			try
			{
				app.DisplayAlerts = false;
				workbook.SaveAs(fileName, XlFileFormat.xlExcel12, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing,
					Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing);
				workbook.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
				app.Quit();
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
		}
		public static void saveFileAsXlsx(string fileName)
		{
			try
			{
				app.DisplayAlerts = false;
				workbook.SaveAs(fileName, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing,
					Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
					Type.Missing, Type.Missing, Type.Missing);
				workbook.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
				app.Quit();
			}
			catch (Exception ex)
			{
				logger.logException(ex);
			}
		}
		public static void setWorkSheetName(string sheetName)
		{
			worksheet.Name = sheetName;
		}
		public static int writeArrayToExcel(int rowPosition, int columnPosition, int rowCount, int columnCount, double[,] info)
		{
			Range initialCell = worksheet.Cells[rowPosition, columnPosition];
			Range lastCell = worksheet.Cells[rowPosition + (rowCount - 1), columnPosition + (columnCount - 1)];
			Range fillingRange = worksheet.Range[initialCell, lastCell];
			fillingRange.Value = info;
			Marshal.ReleaseComObject(initialCell);
			Marshal.ReleaseComObject(lastCell);
			Marshal.ReleaseComObject(fillingRange);
			return columnPosition + columnCount;
		}
		public static int writeArrayToExcel(int rowPosition, int columnPosition, int rowCount, int columnCount, string[,] info)
		{
			Range initialCell = worksheet.Cells[rowPosition, columnPosition];
			Range lastCell = worksheet.Cells[rowPosition + (rowCount - 1), columnPosition + (columnCount - 1)];
			Range fillingRange = worksheet.Range[initialCell, lastCell];
			fillingRange.Value = info;
			Marshal.ReleaseComObject(initialCell);
			Marshal.ReleaseComObject(lastCell);
			Marshal.ReleaseComObject(fillingRange);
			return columnPosition + columnCount;
		}
		public static int writeArrayToExcel(int rowPosition, int columnPosition, int rowCount, int columnCount, object[,] info)
		{
			Range initialCell = worksheet.Cells[rowPosition, columnPosition];
			Range lastCell = worksheet.Cells[rowPosition + (rowCount - 1), columnPosition + (columnCount - 1)];
			Range fillingRange = worksheet.Range[initialCell, lastCell];
			fillingRange.Value = info;
			Marshal.ReleaseComObject(initialCell);
			Marshal.ReleaseComObject(lastCell);
			Marshal.ReleaseComObject(fillingRange);
			return columnPosition + columnCount;
		}
		
		//This function is meant to be used by only the ct air master file since for some reason we cant read it the regular way
		//we have to read it in chunks, in this case we read 65 columns by 1000 rows.
		public static List<object[,]> getRangeListFromExcel(string filePath)
		{
			List<object[,]> rangeList = new List<object[,]>();
			try
			{
				killExcel();
				app = new Application();
				app.DisplayAlerts = false;
				workbook = app.Workbooks.Open(filePath, ReadOnly: true);
				worksheet = workbook.Sheets["Master File"];
				int lowerLimit = 1;
				int upperLimit = 1000;
				bool lastItemTrue = true;

				Thread thread = new Thread(()=>
				{ 
					while(lastItemTrue)
					{
						Range firstCell = worksheet.Cells[lowerLimit, 1];
						Range lastCell = worksheet.Cells[upperLimit, 65];
						Range array = worksheet.Range[firstCell, lastCell];
						if(array[1,1].Value2 == null)
						{
							lastItemTrue = false;
							break;
						}
						upperLimit += 1000;
						lowerLimit += 1000;
						rangeList.Add(array.Value2);
					}
				});
				thread.SetApartmentState(ApartmentState.STA);
				thread.Start();
				thread.Join();
				releaseMemory();
				killExcel();
				return rangeList;
			}
			catch(Exception ex)
			{
				releaseMemory();
				killExcel();
				logger.logException(ex);
				return rangeList;
			}
		}
		public static List<T> ReadInputFiles<T>(string FileName, string SheetName, Dictionary<string, string> FieldToPropertyDictionary) where T : new()
		{
			List<T> itemList = new List<T>();

			Application xlApp = null;
			Workbooks xlWorkBooks = null;
			Workbook xlWorkBook = null;
			Sheets xlWorkSheets = null;
			Worksheet xlWorkSheet = null;
			Range xlRange = null;
			Range xlRangeFormat = null;

			xlApp = new ExcelR.Application();

			xlApp.AskToUpdateLinks = false;
			xlApp.DisplayAlerts = false;
			xlApp.Visible = false;

			xlWorkBooks = xlApp.Workbooks;

			string[] fieldNameList = FieldToPropertyDictionary.Keys.ToArray();
			try
			{
				//create dictionary to map column names to class properties
				Dictionary<string, int> fieldToColumnNumberDictionary = new Dictionary<string, int>();

				//initialize variable which stores the number of the first row after the header row
				int initialDataRow = 1;
				int headerRow = 1;

				xlWorkBook = xlWorkBooks.Open(FileName);
				xlWorkSheets = (Sheets)xlWorkBook.Worksheets;
				xlWorkSheet = (Worksheet)xlWorkSheets[SheetName];

				//get number of used rows and columns
				xlWorkSheet.Columns.ClearFormats();
				xlWorkSheet.Rows.ClearFormats();

				xlWorkSheet.Cells.Replace(What: "|", Replacement: "", LookAt: XlLookAt.xlPart,
											SearchOrder: XlSearchOrder.xlByRows,
											 MatchCase: false, SearchFormat: false, ReplaceFormat: false);
				xlWorkSheet.Cells.Replace(What: "\n", Replacement: "", LookAt: XlLookAt.xlPart,
											SearchOrder: XlSearchOrder.xlByRows,
											MatchCase: false, SearchFormat: false, ReplaceFormat: false);
				xlWorkSheet.Cells.Replace(What: "\r", Replacement: "", LookAt: XlLookAt.xlPart,
											SearchOrder: XlSearchOrder.xlByRows,
											MatchCase: false, SearchFormat: false, ReplaceFormat: false);
				xlWorkSheet.Cells.Replace(What: "\t", Replacement: "", LookAt: XlLookAt.xlPart,
											SearchOrder: XlSearchOrder.xlByRows,
											MatchCase: false, SearchFormat: false, ReplaceFormat: false);

				//get the whole used ranget
				xlRangeFormat = xlWorkSheet.UsedRange;

				//get the entire data table
				var usedRangeValues = (object[,])xlRangeFormat.Value2;

				int nRows = usedRangeValues.GetLength(0);
				int nCols = usedRangeValues.GetLength(1);

				bool hasFoundAnyFieldName = false;
				int columnIndex = 1;
				int rowIndex = 1;

				while (!hasFoundAnyFieldName && columnIndex < (nCols + 1))
				{
					rowIndex = 1;

					while (!hasFoundAnyFieldName && rowIndex < (nRows + 1))
					{
						foreach (var field in fieldNameList)
						{
							if (usedRangeValues[rowIndex, columnIndex] != null)
							{
								if (usedRangeValues[rowIndex, columnIndex].ToString() == field)
								{
									hasFoundAnyFieldName = true;
									headerRow = rowIndex;
									initialDataRow = headerRow + 1;
									break;
								}
							}
						}

						rowIndex++;
					}

					columnIndex++;
				}

				foreach (var field in fieldNameList)
				{
					for (int colIndex = 1; colIndex < (nCols + 1); colIndex++)
					{
						if (usedRangeValues[headerRow, colIndex] != null)
						{
							if (usedRangeValues[headerRow, colIndex].ToString() == field)
							{
								fieldToColumnNumberDictionary.Add(field, colIndex);
								break;
							}
						}
					}
				}

				var columnDictionaryValues = FieldToPropertyDictionary.Values.ToArray();
				var tableRows = (nRows - headerRow + 1);

				for (int rowNumber = initialDataRow; rowNumber < (tableRows + 1); rowNumber++)
				{
					T newItem = new T();

					foreach (var field in fieldNameList)
					{
						string propName = FieldToPropertyDictionary[field];
						var property = newItem.GetType().GetProperty(propName);

						if (property.PropertyType == typeof(DateTime))
						{
							string value =
								usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]] != null
								? usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString()
								: "";

							if (!String.IsNullOrEmpty(value))
							{
								double dateDoubleValue = 0.0;
								bool successConvertingToDouble = double.TryParse(value, out dateDoubleValue);

								if (successConvertingToDouble)
								{
									DateTime date = DateTime.FromOADate(dateDoubleValue);
									property.SetValue(newItem, date);
								}
							}
						}
						else if (property.PropertyType == typeof(int) && usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]] != null)
						{
							int number;
							string debugValue = usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString();
							var valueInt = Int32.TryParse(Regex.Match(usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString(), @"\d+").Value, out number); ;
							if (valueInt)
							{
								if (usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString().Contains("-"))
								{
									number = number * -1;
								}
								property.SetValue(newItem, number);
							}
						}
						else if (property.PropertyType == typeof(float) && usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]] != null)
						{
							float number;
							var valueFloat = float.TryParse(usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString(), out number);

							if (valueFloat)
							{
								property.SetValue(newItem, number);
							}
						}
						else if (property.PropertyType == typeof(double) && usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]] != null)
						{
							float number;
							var valueFloat = float.TryParse(usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]].ToString(), out number);

							if (valueFloat)
							{
								property.SetValue(newItem, number);
							}
						}
						else
						{
							if (usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]] != null)
							{
								var valueToSet = Convert.ChangeType(usedRangeValues[rowNumber, fieldToColumnNumberDictionary[field]], property.PropertyType);
								property.SetValue(newItem, valueToSet);
							}
						}
					}

					itemList.Add(newItem);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.ToString());
				Console.WriteLine(ex.Message);
				System.Console.WriteLine("Error while reading the file.");
			}
			finally
			{
				if (xlRange != null)
				{
					Marshal.FinalReleaseComObject(xlRange);
				}

				if (xlWorkSheet != null)
				{
					Marshal.FinalReleaseComObject(xlWorkSheet);
				}

				if (xlWorkSheets != null)
				{
					Marshal.FinalReleaseComObject(xlWorkSheets);
				}

				if (xlWorkBook != null)
				{
					xlWorkBook.Close(false);
					Marshal.FinalReleaseComObject(xlWorkBook);
				}

				if (xlWorkBooks != null)
				{
					xlWorkBooks.Close();
					Marshal.FinalReleaseComObject(xlWorkBooks);
				}

				if (xlApp != null)
				{
					xlApp.Quit();
					Marshal.FinalReleaseComObject(xlApp);
				}

				GC.Collect();
				GC.WaitForPendingFinalizers();
			}
			return itemList;
		}
	}
}
