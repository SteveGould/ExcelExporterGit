using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using HC = ExcelExport.HelperClasses;
using OpenXMLExtensions;

namespace OpenXMLHelper
{
	public static class OpenXML
	{
		/// <summary>
		/// Creates a new font.  Can be extended as other attributes are needed
		/// </summary>
		/// <param name="fontName"></param>
		/// <param name="fontSize"></param>
		/// <param name="isBold"></param>
		/// <param name="isItalics"></param>
		/// <param name="fontForeColor"></param>
		/// <returns></returns>
		public static Font CreateNewFont(	string fontName, 
											double? fontSize, 
											bool isBold, 
											bool isItalics,
											string fontForeColor
										)
		{
			Font font = new Font();

			if (!string.IsNullOrEmpty(fontName))
			{
				FontName name = new FontName { Val = fontName };
				font.Append(name);
			}

			if (fontSize.HasValue)
			{
				FontSize size = new FontSize { Val = fontSize.Value };
				font.Append(size);
			}

			if (isBold)
			{
				Bold bold = new Bold();
				font.Append(bold);
			}

			if (isItalics)
			{
				Italic italic = new Italic();
				font.Append(italic);
			}

			if (!string.IsNullOrEmpty(fontForeColor))
			{
				Color color = new Color { Rgb = new HexBinaryValue { Value = fontForeColor}};
				font.Append(color);
			}

			return font;
		}

		/// <summary>
		/// Creates a new cellFormat.  Can be extended as other attributes are needed
		/// </summary>
		/// <param name="cell"></param>
		/// <param name="styleSheet"></param>
		/// <returns></returns>
		public static CellFormat CreateCellFormat(HC.Cell cell, Stylesheet styleSheet)
		{
			CellFormat cellFormat = new CellFormat();

			if (cell.Alignment != null)
			{
				cellFormat.ApplyAlignment = true;
				cellFormat.Append(cell.Alignment);
			}

			if (cell.Font != null)
			{
				cellFormat.FontId = cell.Font.GetFontIndex(styleSheet);
				cellFormat.ApplyFont = true;
			}

			if (cell.Fill != null)
			{
				cellFormat.FillId = cell.Fill.GetFillIndex(styleSheet);
				cellFormat.ApplyFill = true;
			}

			if (cell.Border != null)
			{
				cellFormat.BorderId = cell.Border.GetBorderIndex(styleSheet);
				cellFormat.ApplyBorder = true;
			}

			if (cell.NumberingFormat != null)
			{
				if (cell.NumberingFormat.NumberFormatId != null)
					cellFormat.NumberFormatId = cell.NumberingFormat.NumberFormatId;
				else
					cellFormat.NumberFormatId = cell.NumberingFormat.GetNumberingFormatIndex(styleSheet);

				cellFormat.ApplyNumberFormat = true;
			}

			return cellFormat;
		}

		/// <summary>
		/// Creates a new fill.  Can be extended as other attributes are needed
		/// </summary>
		/// <param name="foregroundColor"></param>
		/// <returns></returns>
		public static Fill CreateFill(ForegroundColor foregroundColor)
		{
			ForegroundColor fgc = new ForegroundColor();
			fgc.Theme = foregroundColor.Theme;

			PatternFill patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.Solid;
			patternFill.Append(fgc);

			Fill fill = new Fill(patternFill);

			return fill;
		}

		/// <summary>
		/// Creates a new fill.  Can be extended as other attributes are needed
		/// </summary>
		/// <param name="foregroundColor"></param>
		/// <param name="backgroundColor"></param>
		/// <returns></returns>
		public static Fill CreateFill(ForegroundColor foregroundColor, BackgroundColor backgroundColor)
		{
			PatternFill patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.Solid;
			patternFill.Append(foregroundColor);
			patternFill.Append(backgroundColor);

			Fill fill = new Fill(patternFill);

			return fill;
		}

		/// <summary>
		/// Creates a new Numbering Format.  Can be extended as other attributes are needed
		/// </summary>
		/// <param name="formatCode"></param>
		/// <returns></returns>
		public static NumberingFormat CreateNumberFormat(string formatCode)
		{
			NumberingFormat nf = new NumberingFormat();
			nf.FormatCode = formatCode;

			return nf;
		}

		public static Column CreateColumn(UInt32Value columnNbr, double width)
		{
			Column column = new Column();
			column.Min = columnNbr;
			column.Max = columnNbr;
			column.Width = width;

			column.CustomWidth = true;

			return column;
		}

		public static Column CreateColumnRange(UInt32Value firstColumnNbr, UInt32Value lastColumnNbr, double width)
		{
			Column column = new Column();
			column.Min = firstColumnNbr;
			column.Max = lastColumnNbr;
			column.Width = width;

			column.CustomWidth = true;

			return column;
		}

		/// <summary>
		/// Creates black border around cell
		/// </summary>
		/// <returns></returns>
		public static Border CreateBorder()
		{
			Border border = new Border();

			LeftBorder leftBorder = new LeftBorder() { Style = BorderStyleValues.Thin };
			Color color1 = new Color() { Indexed = (UInt32Value)64U };

			leftBorder.Append(color1);

			RightBorder rightBorder = new RightBorder() { Style = BorderStyleValues.Thin };
			Color color2 = new Color() { Indexed = (UInt32Value)64U };

			rightBorder.Append(color2);

			TopBorder topBorder = new TopBorder() { Style = BorderStyleValues.Thin };
			Color color3 = new Color() { Indexed = (UInt32Value)64U };

			topBorder.Append(color3);

			BottomBorder bottomBorder = new BottomBorder() { Style = BorderStyleValues.Thin };
			Color color4 = new Color() { Indexed = (UInt32Value)64U };

			bottomBorder.Append(color4);

			DiagonalBorder diagonalBorder = new DiagonalBorder();

			border.Append(leftBorder);
			border.Append(rightBorder);
			border.Append(topBorder);
			border.Append(bottomBorder);
			border.Append(diagonalBorder);

			return border;
		}

		public static Pane CreateFreezePane(string topLeftCell)
		{
			Pane pane = new Pane() 
			{ 
				VerticalSplit = 1D, 
				TopLeftCell = topLeftCell, 
				ActivePane = PaneValues.BottomLeft, 
				State = PaneStateValues.Frozen 
			};

			return pane;
		}


	}
}
