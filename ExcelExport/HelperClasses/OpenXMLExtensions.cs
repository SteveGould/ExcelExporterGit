using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXMLExtensions
{
	public static class OpenXMLExtensions
	{
		/// <summary>
		/// Grabs Font index if it exists, otherwise creates a new one and adds it to the stylesheet
		/// </summary>
		/// <param name="font">The font to check</param>
		/// <param name="styleSheet">Stylesheet to check</param>
		/// <returns>The index</returns>
		public static UInt32Value GetFontIndex(this Font font, Stylesheet styleSheet)
		{
			string outerXML = font.OuterXml;
			UInt32Value index = 0;

			foreach (Font f in styleSheet.Fonts.ToList())
			{
				if (f.OuterXml.Equals(outerXML))
					return index;

				index++;
			}

			styleSheet.Fonts.Append(font);

			UInt32Value result = styleSheet.Fonts.Count.Value;
			styleSheet.Fonts.Count++;

			return result;
		}

		/// <summary>
		/// Grabs CellFormat index if it exists, otherwise creates a new one and adds it to the stylesheet
		/// </summary>
		/// <param name="cellFormat">The cellformat to check</param>
		/// <param name="styleSheet">Stylesheet to check</param>
		/// <returns>The index</returns>
		public static UInt32Value GetCellFormatIndex(this CellFormat cellFormat, Stylesheet styleSheet)
		{
			string outerXML = cellFormat.OuterXml;
			UInt32Value index = 0;

			foreach (CellFormat f in styleSheet.CellFormats.ToList())
			{
				if (f.OuterXml.Equals(outerXML))
					return index;

				index++;
			}

			styleSheet.CellFormats.Append(cellFormat);

			UInt32Value result = styleSheet.CellFormats.Count.Value;
			styleSheet.CellFormats.Count++;

			return result;
		}

		/// <summary>
		/// Grabs NumberingFormat index if it exists, otherwise creates a new one and adds it to the stylesheet
		/// </summary>
		/// <param name="numberingFormat">the numberingFormat to check</param>
		/// <param name="styleSheet">Stylesheet to check</param>
		/// <returns>The index</returns>
		public static UInt32Value GetNumberingFormatIndex(this NumberingFormat numberingFormat, Stylesheet styleSheet)
		{
			// adding a custom numbering format does not appear to work.
			// stick with the predefined ones

			string outerXML = numberingFormat.OuterXml;
			bool hasNumFmtID = numberingFormat.NumberFormatId != null;

			UInt32Value index = 0;
			UInt32Value maxNumFmtID = 0;

			foreach (NumberingFormat f in styleSheet.NumberingFormats.ToList())
			{
				if (f.OuterXml.Equals(outerXML))
					return index;

				if (hasNumFmtID && f.NumberFormatId.Value == numberingFormat.NumberFormatId.Value)
					return index;

				if (f.FormatCode.Value == numberingFormat.FormatCode.Value)
					return index;

				maxNumFmtID = Math.Max(maxNumFmtID, f.NumberFormatId);
				index++;
			}

			if (!hasNumFmtID)
				numberingFormat.NumberFormatId = Math.Max(maxNumFmtID, 163U) + 1;

			styleSheet.NumberingFormats.Append(numberingFormat);

			UInt32Value result = styleSheet.NumberingFormats.Count.Value;
			styleSheet.NumberingFormats.Count++;

			return result;
		}

		/// <summary>
		/// Grabs Fill index if it exists, otherwise creates a new one and adds it to the stylesheet
		/// </summary>
		/// <param name="fill">the fill to check</param>
		/// <param name="styleSheet">Stylesheet to check</param>
		/// <returns>The index</returns>
		public static UInt32Value GetFillIndex(this Fill fill, Stylesheet styleSheet)
		{
			string outerXML = fill.OuterXml;
			UInt32Value index = 0;

			foreach (Fill f in styleSheet.Fills.ToList())
			{
				if (f.OuterXml.Equals(outerXML))
					return index;

				index++;
			}

			styleSheet.Fills.Append(fill);

			UInt32Value result = styleSheet.Fills.Count.Value;
			styleSheet.Fills.Count++;

			return result;
		}

		/// <summary>
		/// Grabs Font index if it exists, otherwise creates a new one and adds it to the stylesheet
		/// </summary>
		/// <param name="border">The font to check</param>
		/// <param name="styleSheet">Stylesheet to check</param>
		/// <returns>The index</returns>
		public static UInt32Value GetBorderIndex(this Border border, Stylesheet styleSheet)
		{
			string outerXML = border.OuterXml;
			UInt32Value index = 0;

			foreach (Border f in styleSheet.Borders.ToList())
			{
				if (f.OuterXml.Equals(outerXML))
					return index;

				index++;
			}

			styleSheet.Borders.Append(border);

			UInt32Value result = styleSheet.Borders.Count.Value;
			styleSheet.Borders.Count++;

			return result;
		}
	}
}
