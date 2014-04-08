using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml;

namespace ExcelExporterDemo
{
	public class SpreadSheetAttributes
	{
		public DataTable Data { get; set; }

		// key: Field Name, value: Display Name
		public Dictionary<string, string> HeaderNames { get; set; }

		// key: Field Name, value: Width
		public Dictionary<string, double> FieldWidths { get; set; }
		public Dictionary<string, string> FormatCodes { get; set; }
		public Dictionary<string, UInt32Value> FormatIDs { get; set; }
	}
}