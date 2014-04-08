using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelExport;
using ExcelExport.HelperClasses;
using DOS = DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExporterDemo
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

		public FileResult ExportNoFormatting()
		{
			// do everything in memory
			MemoryStream excelMS;

			TestData data = new TestData();

			// populate with 10,000 rows of data to show how fast it can be
			DataTable dt = data.CreateTestData(10000);

			// pass in the data table to export.  
			excelMS = ExportHelper.GenerateExcel(dt);

			string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			string fileName = "NoFormatting" + ".xlsx";

			// download.  slightly different code if using ASP.NET
			return File(excelMS.ToArray(), contentType, fileName);
		}

		public FileResult ExportWithFormatting()
		{
			// do everything in memory
			MemoryStream excelMS;

			// create a class to hold all the attributes.  just makes it easier
			// to pass it around
			SpreadSheetAttributes ssa = new SpreadSheetAttributes();

			// this is NOT the same worksheet class as in DocumentFormat.OpenXml.Spreadsheet
			// holds formatted data so can pass to the export
			Worksheet ws;

			TestData data = new TestData();

			// set the spread sheet attributes
			ssa.Data = data.CreateTestData(10000);
			ssa.HeaderNames = data.CreateHeaderNames();
			ssa.FormatCodes = data.CreateFormatCodes();
			ssa.FormatIDs = data.CreateFormatIDs();
			ssa.FieldWidths = data.CreateWidths();

			// create all the formatting
			ws = PopulateSpreadsheetAttributes(ssa);

			// run exprot
			excelMS = ExportHelper.GenerateExcel(ws);

			string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
			string fileName = "WithFormatting" + ".xlsx";

			return File(excelMS.ToArray(), contentType, fileName);
		}

		/// <summary>
		/// Populates the worksheet class to set all the formatting
		/// </summary>
		/// <param name="ssa">all the attributes that hold the formatting</param>
		/// <returns>worksheet to be exported</returns>
		private Worksheet PopulateSpreadsheetAttributes(SpreadSheetAttributes ssa)
		{
			Worksheet ws = new Worksheet();

			// set sheet name
			ws.SheetName = "Number Formats";

			// freeze top row.  
			ws.FreezePaneTopLeftCell = "A2";

			List<Cell> columnHeadings = new List<Cell>();
			List<DOS.Column> columns = new List<DOS.Column>();

			// header fill color
			DOS.ForegroundColor fillColor = new DOS.ForegroundColor() { Theme = 2U };

			DocumentFormat.OpenXml.UInt32Value index = 0;
			
			// create the header row formatting
			foreach (KeyValuePair<string, string> fieldName in ssa.HeaderNames)
			{
				Cell cell = new Cell();
				cell.Value = fieldName.Value;
				cell.CellDataType = DOS.CellValues.InlineString;

				// set the font formatting of the header
				cell.Font = OpenXMLHelper.OpenXML.CreateNewFont("", null, true, false, "");

				// set the alignment properties of the header.  
				cell.Alignment = new DOS.Alignment() 
								 {
									Horizontal = DOS.HorizontalAlignmentValues.Center, 
									WrapText=true 
								 };

				// set the color of header.  this shows how to color ANY cell
				cell.Fill = OpenXMLHelper.OpenXML.CreateFill(fillColor);

				// add this to the column header variable so can set the worksheet variable
				columnHeadings.Add(cell);

				// set the column attributes.  basically sets the width
				DOS.Column column = OpenXMLHelper.OpenXML.CreateColumn(index + 1, ssa.FieldWidths[fieldName.Key]);
				columns.Add(column);

				index++;
			}

			// assign the above settings to the worksheet
			ws.ColumnHeadings = columnHeadings;
			ws.Columns = columns;

			// format the data rows
			List<Row> rows = new List<Row>();
			DataTable dt = ssa.Data;
			foreach (DataRow dr in dt.Rows)
			{
				Row row = new Row();
				List<Cell> cells = new List<Cell>();
				foreach (KeyValuePair<string, string> fieldName in ssa.HeaderNames)
				{
					Cell cell = new Cell(fieldName.Key, dr, dt);

					// to set the format string, either of the following ways would work.

					//if (!string.IsNullOrEmpty(ssa.FormatCodes[fieldName.Key]))
					//	cell.CreateNumberingFormat(ssa.FormatCodes[fieldName.Key]);

					if (ssa.FormatIDs[fieldName.Key] != null)
						cell.CreateNumberingFormat(ssa.FormatIDs[fieldName.Key]);

					// create borders around just the string column
					if (fieldName.Key == "String")
					{
						cell.Border = OpenXMLHelper.OpenXML.CreateBorder();
						cell.Alignment = new DOS.Alignment() { Horizontal = DOS.HorizontalAlignmentValues.Right };
					}
					cells.Add(cell);
				}
				row.Cells = cells;
				rows.Add(row);
			}

			ws.Rows = rows;

			return ws;
		}
	}
}
