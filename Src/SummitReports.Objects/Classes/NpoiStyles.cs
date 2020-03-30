using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using Newtonsoft.Json;
using NPOI.SS.Util;
using NPOI.HSSF.Util;

namespace SummitReports.Objects
{
    [Flags]
    public enum CellBorder
    {
        None = 0,
        Bottom = 1,
        Top = 2,
        Left = 4,
        Right = 8,
        TopBottom = 2 + 1,
        LeftRight = 4 + 8,
        LeftBottom = 1 + 4,
        LeftTop = 2 + 4,
        RightBottom = 1 + 8,
        RightTop = 2 + 8,
        All = 1 + 2 + 4 + 8
    }
    public enum FormatStyle
    {
        Default = 0,
        Date = 1,
        Number = 2,
        Decimal = 3,
        Currency = 4,
        Percent = 5
    }

    public static class NpoiColor {
        public static XSSFColor BLACK
        {
            get
            {
                return new XSSFColor(System.Drawing.Color.Black);
            }
        }
        public static XSSFColor WithIndex(this XSSFColor xSSFColor)
        {
            if (xSSFColor == null) return null;
            if (hexStringIndexColorCache.Count() == 0) CacheIndexColors();
            if (hexStringIndexColorCache.ContainsKey("#" + ByteArrayToString(xSSFColor.RGB)))
                xSSFColor.Indexed = hexStringIndexColorCache["#" + ByteArrayToString(xSSFColor.RGB)].Index;
            return xSSFColor;
        }
        /// <summary>
        /// This will change the IColor that shoudld of of type XSSFColor, but also will attempt to find the index color and populate it.  this is important because POI/NPOI cell styles require the index as opposed 
        /// to and XSSFColor object.
        /// </summary>
        /// <param name="colorToConvert"></param>
        /// <returns></returns>
        public static XSSFColor AsXSSFColor(this IColor colorToConvert)
        {
            if (colorToConvert == null) return null;
            return ((XSSFColor)colorToConvert).WithIndex();
        }

        public static XSSFColor AsXSSFColor(this IndexedColors indexedColor)
        {
            if (indexedColor == null) return null;
            return (new XSSFColor(indexedColor.RGB)).WithIndex();
        }
        private static string ByteArrayToString(byte[] ba)
        {
            StringBuilder hex = new StringBuilder(ba.Length * 2);
            foreach (byte b in ba)
                hex.AppendFormat("{0:x2}", b);
            return hex.ToString();
        }

        private static Dictionary<string, IndexedColors> hexStringIndexColorCache = new Dictionary<string, IndexedColors>();
        private static void CacheIndexColors()
        {
            Type type = typeof(IndexedColors); // MyClass is static class with static properties
            foreach (var p in type.GetFields(System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic))
            {
                var colors = (Dictionary<string, IndexedColors>)p.GetValue(null);
                foreach(var color in colors.Keys)
                {
                    var indexColor = colors[color];
                    if (hexStringIndexColorCache.ContainsKey(indexColor.HexString))
                    {
                        Console.WriteLine(indexColor.HexString);
                    } else
                    {
                        hexStringIndexColorCache.Add(indexColor.HexString, indexColor);
                    }
                }
                break;
            }
        }
    }

    public class CellFormats
    {
        public string Default = "";
        public string Date = "mm/dd/yyyy";
        public string Integer = "#,##0";
        public string Decimal = "#,##0.00";
        public string Currency = "$#,##0.00";
        public string Percent = "#0.00%";
        public string FormatStringFromStyle(FormatStyle formatStyle)
        {
            if (formatStyle == FormatStyle.Default)
                return this.Default;
            else if (formatStyle == FormatStyle.Date)
                return this.Date;
            else if (formatStyle == FormatStyle.Currency)
                return this.Currency;
            else if (formatStyle == FormatStyle.Number)
                return this.Integer;
            else if (formatStyle == FormatStyle.Decimal)
                return this.Decimal;
            else if (formatStyle == FormatStyle.Percent)
                return this.Percent;
            return "";
        }
    }
    public class NpoiStyle<T, Y>
    {
        public FormatStyle FormatStyle { get; set; } = FormatStyle.Default;
        public CellBorder Border { get; set; } = CellBorder.None;
        public BorderStyle BorderStyle { get; set; }
        public virtual IColor BorderColor { get; set; } = null;
        public string CellFormat {
            get {
                return this.CellFormats.Default;
            }
            set {
                this.CellFormats.Default = value;
            }
        }
        public virtual Y SetFormatStyle(FormatStyle style) {
            throw new Exception("Must override this function!!!");
        }
        public CellFormats CellFormats { get; set; } = new CellFormats();
        public double? FontHeightInPoints { get; set; } = null;
        public string FontName { get; set; } = null;
        public virtual IColor FontColor { get; set; } = null;
        public bool? IsBold { get; set; } = null;
        public bool? IsItalic { get; set; } = null;
        public HorizontalAlignment? HorizontalAlignment { get; set; } = null;
        public VerticalAlignment? VerticalAlignment { get; set; } = null;
        public virtual IColor BackgroundColor { get; set; } = null;
        public virtual IColor FillForegroundColor { get; set; } = null;
        public FillPattern? FillPattern { get; set; } = null;
        public bool? WrapText { get; set; } = null;
        public NpoiStyle()
        {
        }
        protected virtual T Render(IWorkbook workbook, FormatStyle formatStyle) {
            throw new Exception("Must override this function!!!");
        }
        private bool hasRendered = false;
        public T Render(ICell cell, FormatStyle formatStyle) {
            return Render(cell.Sheet.Workbook, formatStyle);
        }

        public T Render(ICell cell)
        {
            if (!this.hasRendered)
            {
                foreach (FormatStyle formatStyle in (FormatStyle[])Enum.GetValues(typeof(FormatStyle)))
                {
                    Render(cell.Sheet.Workbook, formatStyle);
                }
                this.hasRendered = true;
            }
            return Render(cell.Sheet.Workbook, this.FormatStyle);
        }
    }

    public class XSSFNPoiStyle : NpoiStyle<XSSFCellStyle, XSSFNPoiStyle>
    {
        public override XSSFNPoiStyle SetFormatStyle(FormatStyle style)
        {
            this.FormatStyle = style;
            this.CellFormat = this.CellFormats.FormatStringFromStyle(style);
            return this;
        }
        public XSSFNPoiStyle SetFormatStyle(string customFormatString)
        {
            this.FormatStyle = FormatStyle.Default;
            this.CellFormat = customFormatString;
            return this;
        }

        private string GetSha256Hash(string input)
        {

            using (SHA256 hash = SHA256Managed.Create())
            {
                return String.Concat(hash
                  .ComputeHash(Encoding.UTF8.GetBytes(input))
                  .Select(item => item.ToString("x2")));
            }
            // Convert the input string to a byte array and compute the hash.
        }
        public CellRangeAddress ApplyBorderToRange(ISheet worksheet, CellRangeAddress range)
        {
            var defaultCellStyle = worksheet.Workbook.GetCellStyleAt(0);
            if (this.Border != CellBorder.None)
            {
                if (this.Border.HasFlag(CellBorder.Top))
                {
                    RegionUtil.SetBorderTop((int)this.BorderStyle, range, worksheet, worksheet.Workbook);
                    if (this.BorderColor == null)
                    {
                        RegionUtil.SetTopBorderColor(defaultCellStyle.TopBorderColor, range, worksheet, worksheet.Workbook);
                    }
                    else
                    {
                        throw new Exception("RegionUtil.SetTopBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook)");
                        //RegionUtil.SetTopBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook);
                    }
                }
                if (this.Border.HasFlag(CellBorder.Bottom))
                {
                    RegionUtil.SetBorderBottom((int)this.BorderStyle, range, worksheet, worksheet.Workbook);
                    if (this.BorderColor == null)
                    {
                        RegionUtil.SetBottomBorderColor(defaultCellStyle.BottomBorderColor, range, worksheet, worksheet.Workbook);
                    }
                    else
                    {
                        throw new Exception("RegionUtil.SetBottomBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook)");
                        //RegionUtil.SetBottomBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook);
                    }
                }
                if (this.Border.HasFlag(CellBorder.Right))
                {
                    RegionUtil.SetBorderRight((int)this.BorderStyle, range, worksheet, worksheet.Workbook);
                    if (this.BorderColor == null)
                    {
                        RegionUtil.SetRightBorderColor(defaultCellStyle.RightBorderColor, range, worksheet, worksheet.Workbook);
                    }
                    else
                    {
                        throw new Exception("RegionUtil.SetRightBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook)");
                        //RegionUtil.SetRightBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook);
                    }
                }
                if (this.Border.HasFlag(CellBorder.Left))
                {
                    RegionUtil.SetBorderLeft((int)this.BorderStyle, range, worksheet, worksheet.Workbook);
                    if (this.BorderColor == null)
                    {
                        RegionUtil.SetLeftBorderColor(defaultCellStyle.LeftBorderColor, range, worksheet, worksheet.Workbook);
                    }
                    else
                    {
                        throw new Exception("egionUtil.SetLeftBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook)");
                        //RegionUtil.SetLeftBorderColor(this.BorderColor.Indexed, range, worksheet, worksheet.Workbook);
                    }
                }
            }
            return range;
        }
        protected override XSSFCellStyle Render(IWorkbook workbook, FormatStyle formatStyle)
        {
            //var saveCellFormat = this.CellFormat;
            //this.CellFormat = this.CellFormats.FormatStringFromStyle(formatStyle);
            string thisStyleCacheKey = GetSha256Hash(JsonConvert.SerializeObject(this));
            //this.CellFormat = saveCellFormat;

            if (NPoiExtentions.StyleCache.ContainsKey(thisStyleCacheKey))
            {
                return (XSSFCellStyle)NPoiExtentions.StyleCache[thisStyleCacheKey];
            }

            XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            var defaultCellStyle = workbook.GetCellStyleAt(0);
            if (this.Border != CellBorder.None)
            {
                if (this.Border.HasFlag(CellBorder.Top))
                {
                    cellStyle.BorderTop = this.BorderStyle;
                    if (this.BorderColor==null)
                    {
                        cellStyle.TopBorderColor = defaultCellStyle.TopBorderColor;
                    } else
                    {
                        cellStyle.SetTopBorderColor(this.BorderColor.AsXSSFColor());
                        cellStyle.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.TOP, this.BorderColor.AsXSSFColor());
                    }
                }
                if (this.Border.HasFlag(CellBorder.Bottom))
                {
                    cellStyle.BorderBottom = this.BorderStyle;
                    if (this.BorderColor == null)
                    {
                        cellStyle.BottomBorderColor = defaultCellStyle.BottomBorderColor;
                    }
                    else
                    {
                        cellStyle.SetBottomBorderColor(this.BorderColor.AsXSSFColor());
                        cellStyle.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.BOTTOM, this.BorderColor.AsXSSFColor());
                    }
                }
                if (this.Border.HasFlag(CellBorder.Right))
                {
                    cellStyle.BorderRight = this.BorderStyle;
                    if (this.BorderColor == null)
                    {
                        cellStyle.RightBorderColor = defaultCellStyle.RightBorderColor;
                    }
                    else
                    {
                        cellStyle.SetRightBorderColor(this.BorderColor.AsXSSFColor());
                        cellStyle.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.RIGHT, this.BorderColor.AsXSSFColor());
                    }
                }
                if (this.Border.HasFlag(CellBorder.Left))
                {
                    cellStyle.BorderLeft = this.BorderStyle;
                    if (this.BorderColor == null)
                    {
                        cellStyle.LeftBorderColor = defaultCellStyle.LeftBorderColor;
                    }
                    else
                    {
                        cellStyle.SetLeftBorderColor(this.BorderColor.AsXSSFColor());
                        cellStyle.SetBorderColor(NPOI.XSSF.UserModel.Extensions.BorderSide.LEFT, this.BorderColor.AsXSSFColor());
                    }
                }
            }
            if ((this.FontHeightInPoints.HasValue) || (this.FontName != null) || (this.FontColor != null) || (this.IsBold.HasValue) || (this.IsItalic.HasValue))
            {
                var defaultFont = workbook.GetFontAt(0);
                var font = workbook.CreateFont();
                font.FontHeightInPoints = (this.FontHeightInPoints.HasValue) ? this.FontHeightInPoints.Value : defaultFont.FontHeightInPoints;
                font.FontName = (this.FontName!=null) ? this.FontName : defaultFont.FontName;  // this.FontName;
                font.Color = (this.FontColor!=null) ? this.FontColor.AsXSSFColor().Indexed: defaultFont.Color;  //this.FontColor.Indexed;
                font.IsBold = (this.IsBold.HasValue) ? this.IsBold.Value : defaultFont.IsBold;  //this.IsBold;
                font.IsItalic = (this.IsItalic.HasValue) ? this.IsItalic.Value : defaultFont.IsItalic;  //this.IsItalic;
                cellStyle.SetFont(font);
            }
            if (this.BackgroundColor != null) cellStyle.SetFillBackgroundColor(this.BackgroundColor.AsXSSFColor());
            if (this.HorizontalAlignment != null) cellStyle.Alignment = this.HorizontalAlignment.Value;
            if (this.VerticalAlignment != null) cellStyle.VerticalAlignment = this.VerticalAlignment.Value;
            if (!string.IsNullOrEmpty(this.CellFormat)) cellStyle.DataFormat = workbook.CreateDataFormat().GetFormat(this.CellFormat);
            if (this.FillForegroundColor != null) cellStyle.SetFillForegroundColor(FillForegroundColor.AsXSSFColor());
            if (this.FillPattern != null) cellStyle.FillPattern = this.FillPattern.Value;
            if (this.WrapText != null) cellStyle.WrapText = this.WrapText.Value;
            

            if (!NPoiExtentions.StyleCache.ContainsKey(thisStyleCacheKey)) NPoiExtentions.StyleCache.Remove(thisStyleCacheKey);
            NPoiExtentions.StyleCache.Add(thisStyleCacheKey, cellStyle);
            return cellStyle;
        }
    }
}
