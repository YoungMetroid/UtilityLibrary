using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;
namespace UtilityLibrary.Excel
{
	public class CsvReader
	{
		DataTable csvData;
		string filePath;
		Logger logger = Logger.getInstance;
		public CsvReader(string filePath)
		{
			this.filePath = filePath;
			csvData = new DataTable();
		}
		public void setDataTable()
		{
			try
			{
				using (TextFieldParser csvReader = new TextFieldParser(filePath))
				{
					csvReader.SetDelimiters(new string[] { "," });
					string[] headers = csvReader.ReadFields();
					foreach(string header in headers)
					{
						DataColumn dataColumn = new DataColumn(header);
						dataColumn.AllowDBNull = true;
						csvData.Columns.Add(dataColumn);
					}
					while(!csvReader.EndOfData)
					{
						string[] dataRow = csvReader.ReadFields();

						for(int columnCounter = 0; columnCounter < dataRow.Length; columnCounter++)
						{
							if (dataRow[columnCounter] == "") dataRow[columnCounter] = null;
						}
						csvData.Rows.Add(dataRow);
					}
				}
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
		}
		public DataTable getDataTable()
		{
			return csvData;
		}

	}
}
