using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using OX = DocumentFormat.OpenXml;

namespace ExcelExporterDemo
{
	public class TestData
	{
		public DataTable CreateTestData(int numberOfRows)
		{
			DataTable dataTable = new DataTable("Test");

			CreateColumns(dataTable);

			CreateRows(dataTable, numberOfRows);

			return dataTable;
		}

		public Dictionary<string, string> CreateHeaderNames()
		{
			Dictionary<string, string> headerNames = new Dictionary<string, string>();
		
			headerNames.Add("Date", "Short\nDate");
			headerNames.Add("Integer", "Whole\nNumber");
			headerNames.Add("Float", "Floating\nPoint");
			headerNames.Add("Currency", "Dollars");
			headerNames.Add("Percent", "Percentage");
			headerNames.Add("String", "GUID");

			return headerNames;
		}

		public Dictionary<string, OX.UInt32Value> CreateFormatIDs()
		{
			Dictionary<string, OX.UInt32Value> formats = new Dictionary<string, OX.UInt32Value>();

			formats.Add("Date", 14);
			formats.Add("Integer", 1);
			formats.Add("Float", 2);
			formats.Add("Currency", 8);
			formats.Add("Percent", 10);
			formats.Add("String", null);

			return formats;
		}

		public Dictionary<string, string> CreateFormatCodes()
		{
			Dictionary<string, string> formats = new Dictionary<string, string>();

			formats.Add("Date", "mm/dd/yy");
			formats.Add("Integer", "0");
			formats.Add("Float", "0.00");
			formats.Add("Currency", "$#,##0.00_);[Red]($#,##0.00)");
			formats.Add("Percent", "0.00%");
			formats.Add("String", "");

			return formats;
		}

		public Dictionary<string, double> CreateWidths()
		{
			Dictionary<string, double> formats = new Dictionary<string, double>();

			formats.Add("Date", 10.5);
			formats.Add("Integer", 9.8);
			formats.Add("Float", 10.8);
			formats.Add("Currency", 9.0);
			formats.Add("Percent", 12.2);
			formats.Add("String", 42);

			return formats;
		}

		private void CreateRows(DataTable dataTable, int numberOfRows)
		{
			DateTime dt = new DateTime();
			int i;
			float f;
			decimal c;
			double p;
			string s;

			Random random = new Random();

			for (int k = 0; k < numberOfRows; k++)
			{
				i = random.Next(10000);
				f = (float)random.NextDouble() * (float)random.Next(1, 1000);
				c = (decimal)random.NextDouble() * (decimal)random.Next(1, 1000);
				p = random.NextDouble() * random.Next(1, 1000);
				s = Guid.NewGuid().ToString();
				dt = RandomDay(random);

				dataTable.Rows.Add(dt, i, f, c, p, s);
			}
		}

		private void CreateColumns(DataTable dataTable)
		{
			dataTable.Columns.Add("Date",		typeof(DateTime));
			dataTable.Columns.Add("Integer",	typeof(int));
			dataTable.Columns.Add("Float",		typeof(float));
			dataTable.Columns.Add("Currency",	typeof(decimal));
			dataTable.Columns.Add("Percent",	typeof(double));
			dataTable.Columns.Add("String",		typeof(string));
		}

		private DateTime RandomDay(Random random)
		{
			DateTime start = new DateTime(1960, 1, 1);

			int range = (DateTime.Today - start).Days;

			return start.AddDays(random.Next(range));
		}
	}
}