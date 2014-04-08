using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OX = DocumentFormat.OpenXml;
using DOS = DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExport.HelperClasses
{
	/// <summary>
	/// Holds attributes so can format the spreadsheet
	/// </summary>
	public class Worksheet
	{
		public string SheetName;
		public string FreezePaneTopLeftCell;
		public IEnumerable<DOS.Column> Columns;
		public IEnumerable<Cell> ColumnHeadings;
		public IEnumerable<Row> Rows;
	}

	/// <summary>
	/// all the rows of the spreadsheet
	/// </summary>
	public class Row
	{
		public IEnumerable<Cell> Cells;
	}

	/// <summary>
	/// formatted cell which is passed to exporter.  
	/// uses the same classes as DocumentFormat.OpenXml.Spreadsheet for
	/// ease of creation in the exporter
	/// </summary>
	public class Cell
	{
		// Standard formats in Excel
		public Dictionary<string, OX.UInt32Value> StandardFormats = new Dictionary<string, OX.UInt32Value>
        {
            { "0",								1   },
            { "0.00",							2   },
            { "#,##0",							3   },
			{ "#,##0.00",						4   },
			{ "$#,##0_);($#,##0)",				5   },
			{ "$#,##0_);[Red]($#,##0)",			6   },
			{ "$#,##0.00_);($#,##0.00)",		7   },
			{ "$#,##0.00_);[Red]($#,##0.00)",	8   },
			{ "0%",								9   },
            { "0.00%",							10  },
            { "0.00E+00",						11  },
            { "# ?/?",							12  },
            { "# ??/??",						13  },
            { "mm/dd/yy",						14  },
            { "d-mmm-yy",						15  },
            { "d-mmm",							16  },
            { "mmm-yy",							17  },
            { "h:mm AM/PM",						18  },
            { "h:mm:ss AM/PM",					19  },
            { "h:mm",							20  },
            { "h:mm:ss",						21  },
            { "h/d/yy h:mm",					22  },
            { "#,##0;(#,##0)",					37  },
            { "#,##0;[Red](#,##0)",				38  },
            { "#,##0.00;(#,##0.00)",			39  },
            { "#,##0.00;[Red](#,##0.00)",		40  },
            { "mm:ss",							45  },
            { "[h]:mm:ss",						46  },
            { "mmss.0",							47  },
            { "##0.0E+0",						48  },
            { "@",								49  },
        };

		public string Value;
		public DOS.CellValues CellDataType;
		public DOS.Font Font;
		public DOS.Fill Fill;
		public DOS.Border Border;
		public DOS.NumberingFormat NumberingFormat;
		public DOS.Alignment Alignment;

		public Cell()
		{
		}

		/// <summary>
		/// Create the cell
		/// </summary>
		/// <param name="dc">data column used to determine data type</param>
		/// <param name="dr">data row.  holds value to display</param>
		public Cell(DataColumn dc, DataRow dr)
		{
			Value = dr[dc.ColumnName].ToString();

			AssignType(Type.GetTypeCode(dc.DataType));
		}

		/// <summary>
		/// Create the cell
		/// </summary>
		/// <param name="fieldName">fieldname in row to display</param>
		/// <param name="dr">data row.  holds value to display</param>
		/// <param name="dt">data table.  used to get data type of column</param>
		public Cell(string fieldName, DataRow dr, DataTable dt)
		{
			DataColumn dc = dt.Columns[fieldName];
			Value = dr[fieldName].ToString();

			AssignType(Type.GetTypeCode(dc.DataType));
		}

		/// <summary>
		/// Create the cell
		/// </summary>
		/// <param name="v">object to display.  </param>
		public Cell(object v)
		{
			Type type;

			if (v == null)
			{
				Value = "";
				type = String.Empty.GetType();
			}
			else
			{
				Value = v.ToString();
				type = v.GetType();
			}

			AssignType(Type.GetTypeCode(type));
		}

		/// <summary>
		/// Assigns the type to the data cell
		/// </summary>
		/// <param name="tc">type code of data column</param>
		private void AssignType(TypeCode tc)
		{
			switch (tc)
			{
				case TypeCode.Int16:
				case TypeCode.Int32:
				case TypeCode.Int64:
				case TypeCode.Single:
				case TypeCode.Double:
				case TypeCode.Decimal:
					CellDataType = DOS.CellValues.Number;
					break;

				case TypeCode.String:
				case TypeCode.Char:
					CellDataType = DOS.CellValues.InlineString;
					break;

				case TypeCode.Boolean:
					CellDataType = DOS.CellValues.Boolean;
					break;

				case TypeCode.DateTime:
					CellDataType = DOS.CellValues.Date;
					NumberingFormat = new DOS.NumberingFormat() { NumberFormatId = 14U };	// see above format codes

					if (!string.IsNullOrEmpty(Value))
					{
						// strip off time
						DateTime dateTime = DateTime.Parse(Value).Date;

						// convert it to a format Excel understands
						Value = dateTime.ToOADate().ToString();
					}

					break;

				default:
					throw new Exception("Unknown data type");
					break;
			}
		}

		/// <summary>
		/// Determines if the cell has any formatting associated with it
		/// </summary>
		/// <returns>True: has formatting set somewhere</returns>
		public bool HasFormatting()
		{
			if (Alignment != null)
				return true;

			if (Font != null)
				return true;

			if (Fill != null)
				return true;

			if (Border != null)
				return true;

			if (NumberingFormat != null)
				return true;

			return false;
		}

		/// <summary>
		/// Creates the number formatting class
		/// </summary>
		/// <param name="format">the format code to use</param>
		public void CreateNumberingFormat(string format)
		{
			NumberingFormat = new DOS.NumberingFormat();

			if (StandardFormats.ContainsKey(format))
				NumberingFormat.NumberFormatId = StandardFormats[format];
			else
				NumberingFormat.FormatCode = format;
		}

		/// <summary>
		/// Creates the number formatting class
		/// </summary>
		/// <param name="formatID">the format id to use</param>
		public void CreateNumberingFormat(OX.UInt32Value formatID)
		{
			NumberingFormat = new DOS.NumberingFormat() { NumberFormatId = formatID };
		}
	}
}
