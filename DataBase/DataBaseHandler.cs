using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UtilityLibrary.Loggers;

namespace UtilityLibrary.DataBase
{
	
	public class DataBaseHandler
	{
		private Logger logger;
		private ConsoleLogger consoleLogger;
		private SqlConnection connection;
		SqlCommand sqlStatement;
		SqlDataReader reader;
		private string currentStatement;
		public string _ConnectionString { get; set; }
		public DataBaseHandler()
		{
			logger = Logger.getInstance;
			consoleLogger = ConsoleLogger.getInstance;
		}
		public void setup()
		{

		}
		public bool OpenConnection()
		{
			try
			{
				connection = new SqlConnection(_ConnectionString);
				connection.Open();
				return true;
			}
			catch (Exception ex)
			{
				consoleLogger.logError("Unable to setup a connection with the following connection: \n" +
					_ConnectionString + " this could be a issue with the VPN, Network, etc");
				logger.logException(ex);
			}
			return false;
		}
		public bool CloseConnection()
		{
			try
			{
				connection.Close();
			}
			catch (Exception ex)
			{
				consoleLogger.logError("Unable to close the connection with the following connection: \n" +
					_ConnectionString + " this could be a issue with the VPN, Network, etc or you already \n" +
					"Closed the connection");
				logger.logException(ex);
			}
			return false;
		}
		public void SetSQLQuery(string query)
		{
			if (query == null || query == string.Empty)
			{
				consoleLogger.logError("SetSQLStatement(): The provided SQL Statement is null or invalid: \n" + query);
				return;
			}

			SqlCommand auxCommand = null;

			try
			{
				auxCommand = new SqlCommand(query);
			}
			catch (Exception ex)
			{
				consoleLogger.logError("SetSQLStatement(): The provided SQL Statement '" + query +
					"' couldn't be interpreted by the SQL parser. Please, try again. Details below:\n" +
					ex.Message);
				logger.logException(ex);
			}

			if (auxCommand != null)
			{
				// Success in trying to parse the provided input as a SQL Statement.
				sqlStatement = auxCommand;
				currentStatement = query;
				sqlStatement.Connection = connection;
			}
		}

		public List<List<object>> executeQuery()
		{
			if (sqlStatement == null)
			{
				consoleLogger.logImportantMessage("ExecuteQuery(): There is no SQL statement " +
					"currently pre-loaded in this DatabaseServer instance.");

			}
			if (sqlStatement.Connection == null || sqlStatement.Connection.ToString() == string.Empty)
			{
				consoleLogger.logImportantMessage("ExecuteQuery(): There is no connection " +
					"associated to the current SQL Statement. We will setup it up if available.");
				sqlStatement.Connection = connection;
			}
			List<List<object>> resultTable = new List<List<object>>();
			try
			{
				reader = this.sqlStatement.ExecuteReader();

				int lastColumn = reader.FieldCount;

				while (reader.Read())
				{
					List<object> row = new List<object>();
					for (int columnCounter = 0; columnCounter < lastColumn; columnCounter++)
					{
						if (reader.GetValue(columnCounter) == null ||
							reader.GetValue(columnCounter).ToString() == string.Empty)

							row.Add(null);
						else
						{
							row.Add(reader.GetValue(columnCounter));
						}
					}
					resultTable.Add(row);
				}
				reader.Close();
			}
			catch (SqlException ex)
			{
				consoleLogger.logError("There was an error when trying to execute the following query + \n" +
				currentStatement + " \n" + ex.Message);
				logger.logException(ex);
			}
			catch (Exception ex)
			{
				consoleLogger.logError("There was an error when trying to execute the following query + \n" +
				currentStatement + " \n" + ex.Message);
				logger.logException(ex);
			}
			return resultTable;
		}
		public bool execute()
		{
			if (sqlStatement == null)
			{
				consoleLogger.logImportantMessage("ExecuteQuery(): There is no SQL statement " +
					"currently pre-loaded in this DatabaseServer instance.");
				return false;

			}
			if (sqlStatement.Connection == null || sqlStatement.Connection.ToString() == string.Empty)
			{
				consoleLogger.logImportantMessage("ExecuteQuery(): There is no connection " +
					"associated to the current SQL Statement. We will setup it up if available.");
				sqlStatement.Connection = connection;
			}
			List<List<object>> resultTable = new List<List<object>>();
			try
			{
				this.sqlStatement.ExecuteNonQuery();
				return true;
			}
			catch (SqlException ex)
			{
				consoleLogger.logError("There was an error when trying to execute the following query + \n" +
				currentStatement + " \n" + ex.Message);
				logger.logException(ex);
			}
			catch (Exception ex)
			{
				consoleLogger.logError("There was an error when trying to execute the following query + \n" +
				currentStatement + " \n" + ex.Message);
				logger.logException(ex);
			}
			return false;
		}
		public void resetStatement()
		{
			currentStatement = "";
			sqlStatement = null;
		}
	}
}
