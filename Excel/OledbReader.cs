using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace UtilityLibrary.Excel
{
	public class OledbReader
	{
		private string fileName;
		private string stringConnection;
		private string sheetName;
		private OleDbConnection connection;
		private Logger logger;
		public OledbReader()
		{
			logger = Logger.getInstance;
		}
		public void setFile(string fileName)
		{
			this.fileName = fileName;
		}

		public void setConnection()
		{
			stringConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
				 this.fileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=Yes;IMEX=1;MAXSCANROWS=0'";
		}
		public void setSheet(string sheetName)
		{
			this.sheetName = sheetName;
		}
		public void openConnection()
		{
			connection = new OleDbConnection(stringConnection);
		}
		public void closeConnection()
		{
			try
			{
				connection.Close();
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
		}
		public DataTable getTable()
		{
			DataTable dataTable = new DataTable();
			try
			{
				using (OleDbCommand cmd = new OleDbCommand())
				{
					cmd.Connection = connection;
					cmd.CommandType = CommandType.Text;
					cmd.CommandText = $"SELECT * FROM [{sheetName + '$'}]";
					using (OleDbDataAdapter oleda = new OleDbDataAdapter(cmd))
					{
						oleda.Fill(dataTable);
					}
				}
			}
			catch(Exception ex)
			{
				logger.logException(ex);
			}
			return dataTable;
		}
	}
}
