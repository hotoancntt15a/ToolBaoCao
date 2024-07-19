using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace zModules.Export
{
    public static class Convert
    {
        public static HSSFWorkbook toXLS(this XSSFWorkbook source)
        {
            //Install-Package NPOI -Version 2.0.6
            HSSFWorkbook retVal = new HSSFWorkbook();
            for (int i = 0; i < source.NumberOfSheets; i++)
            {
                HSSFSheet hssfSheet = (HSSFSheet)retVal.CreateSheet(source.GetSheetAt(i).SheetName);
                XSSFSheet xssfsheet = (XSSFSheet)source.GetSheetAt(i);
                CopySheets(xssfsheet, hssfSheet, retVal);
            }
            return retVal;
        } 

        public static XSSFWorkbook toXLSX(this HSSFWorkbook source)
        {
            var destination = new XSSFWorkbook();
            for (int i = 0; i < source.NumberOfSheets; i++)
            {
                var xssfSheet = (XSSFSheet)destination.CreateSheet(source.GetSheetAt(i).SheetName);
                var hssfSheet = (HSSFSheet)source.GetSheetAt(i);
                CopyStyles(hssfSheet, xssfSheet);
                CopySheets(hssfSheet, xssfSheet);
            }
            return destination;
        }

        private static void CopyStyles(HSSFSheet from, XSSFSheet to)
        {
            for (short i = 0; i <= from.Workbook.NumberOfFonts; i++) { CopyFont(to.Workbook.CreateFont(), from.Workbook.GetFontAt(i)); }
            for (short i = 0; i < from.Workbook.NumCellStyles; i++) { CopyStyle(to.Workbook.CreateCellStyle(), from.Workbook.GetCellStyleAt(i), to.Workbook, from.Workbook); }
        }

        private static void CopyFont(IFont toFront, IFont fontFrom)
        {
            toFront.Charset = fontFrom.Charset;
            toFront.Color = fontFrom.Color;
            toFront.FontHeightInPoints = fontFrom.FontHeightInPoints;
            toFront.FontName = fontFrom.FontName;
            toFront.IsBold = fontFrom.IsBold;
            toFront.IsItalic = fontFrom.IsItalic;
            toFront.IsStrikeout = fontFrom.IsStrikeout;
            //toFront.Underline = fontFrom.Underline; <- bug in npoi setter
        }

        private static void CopyStyle(ICellStyle toCellStyle, ICellStyle fromCellStyle, IWorkbook toWorkbook, IWorkbook fromWorkbook)
        {
            toCellStyle.Alignment = fromCellStyle.Alignment;
            toCellStyle.BorderBottom = fromCellStyle.BorderBottom;
            toCellStyle.BorderDiagonal = fromCellStyle.BorderDiagonal;
            toCellStyle.BorderDiagonalColor = fromCellStyle.BorderDiagonalColor;
            toCellStyle.BorderDiagonalLineStyle = fromCellStyle.BorderDiagonalLineStyle;
            toCellStyle.BorderLeft = fromCellStyle.BorderLeft;
            toCellStyle.BorderRight = fromCellStyle.BorderRight;
            toCellStyle.BorderTop = fromCellStyle.BorderTop;
            toCellStyle.BottomBorderColor = fromCellStyle.BottomBorderColor;
            toCellStyle.DataFormat = fromCellStyle.DataFormat;
            toCellStyle.FillBackgroundColor = fromCellStyle.FillBackgroundColor;
            toCellStyle.FillForegroundColor = fromCellStyle.FillForegroundColor;
            toCellStyle.FillPattern = fromCellStyle.FillPattern;
            toCellStyle.Indention = fromCellStyle.Indention;
            toCellStyle.IsHidden = fromCellStyle.IsHidden;
            toCellStyle.IsLocked = fromCellStyle.IsLocked;
            toCellStyle.LeftBorderColor = fromCellStyle.LeftBorderColor;
            toCellStyle.RightBorderColor = fromCellStyle.RightBorderColor;
            toCellStyle.Rotation = fromCellStyle.Rotation;
            toCellStyle.ShrinkToFit = fromCellStyle.ShrinkToFit;
            toCellStyle.TopBorderColor = fromCellStyle.TopBorderColor;
            toCellStyle.VerticalAlignment = fromCellStyle.VerticalAlignment;
            toCellStyle.WrapText = fromCellStyle.WrapText;
            toCellStyle.SetFont(toWorkbook.GetFontAt((short)(fromCellStyle.GetFont(fromWorkbook).Index + 1)));
        }

        private static void CopySheets(HSSFSheet source, XSSFSheet destination)
        {
            var maxColumnNum = 0;
            var mergedRegions = new List<CellRangeAddress>();
            var styleMap = new Dictionary<int, HSSFCellStyle>();
            for (int i = source.FirstRowNum; i <= source.LastRowNum; i++)
            {
                var srcRow = (HSSFRow)source.GetRow(i);
                var destRow = (XSSFRow)destination.CreateRow(i);
                if (srcRow != null)
                {
                    CopyRow(source, destination, srcRow, destRow, mergedRegions);
                    if (srcRow.LastCellNum > maxColumnNum) { maxColumnNum = srcRow.LastCellNum; }
                }
            }
            for (int i = 0; i <= maxColumnNum; i++) { destination.SetColumnWidth(i, source.GetColumnWidth(i)); }
        }

        private static void CopySheets(XSSFSheet source, HSSFSheet destination, HSSFWorkbook retVal)
        {
            int maxColumnNum = 0;
            Dictionary<int, XSSFCellStyle> styleMap = new Dictionary<int, XSSFCellStyle>();
            for (int i = source.FirstRowNum; i <= source.LastRowNum; i++)
            {
                XSSFRow srcRow = (XSSFRow)source.GetRow(i);
                HSSFRow destRow = (HSSFRow)destination.CreateRow(i);
                if (srcRow != null)
                {
                    CopyRow(source, destination, srcRow, destRow, styleMap, retVal);
                    if (srcRow.LastCellNum > maxColumnNum) { maxColumnNum = srcRow.LastCellNum; }
                }
            }
            for (int i = 0; i <= maxColumnNum; i++) { destination.SetColumnWidth(i, source.GetColumnWidth(i)); }
        }
        private static void CopyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow, List<CellRangeAddress> mergedRegions)
        {
            destRow.Height = srcRow.Height;
            for (int j = srcRow.FirstCellNum; srcRow.LastCellNum >= 0 && j <= srcRow.LastCellNum; j++)
            {
                var oldCell = (HSSFCell)srcRow.GetCell(j);
                var newCell = (XSSFCell)destRow.GetCell(j);
                if (oldCell != null)
                {
                    if (newCell == null) { newCell = (XSSFCell)destRow.CreateCell(j); }
                    CopyCell(oldCell, newCell);
                    var mergedRegion = GetMergedRegion(srcSheet, srcRow.RowNum, (short)oldCell.ColumnIndex);
                    if (mergedRegion != null)
                    {
                        var newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow, mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);
                        if (IsNewMergedRegion(newMergedRegion, mergedRegions))
                        {
                            mergedRegions.Add(newMergedRegion);
                            destSheet.AddMergedRegion(newMergedRegion);
                        }
                    }
                }
            }
        }

        private static void CopyCell(HSSFCell oldCell, XSSFCell newCell)
        {
            CopyCellStyle(oldCell, newCell); CopyCellValue(oldCell, newCell);
        }

        private static void CopyCellValue(HSSFCell oldCell, XSSFCell newCell)
        {
            switch (oldCell.CellType)
            {
                case CellType.String:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;

                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;

                case CellType.Blank:
                    newCell.SetCellType(CellType.Blank);
                    break;

                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;

                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;

                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;

                default:
                    break;
            }
        }

        private static void CopyCellStyle(HSSFCell oldCell, XSSFCell newCell)
        {
            if (oldCell.CellStyle == null) return;
            newCell.CellStyle = newCell.Sheet.Workbook.GetCellStyleAt((short)(oldCell.CellStyle.Index + 1));
        }

        private static CellRangeAddress GetMergedRegion(HSSFSheet sheet, int rowNum, short cellNum)
        {
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var merged = sheet.GetMergedRegion(i);
                if (merged.IsInRange(rowNum, cellNum)) { return merged; }
            }
            return null;
        }

        private static void CopyRow(XSSFSheet srcSheet, HSSFSheet destSheet, XSSFRow srcRow, HSSFRow destRow,
                Dictionary<int, XSSFCellStyle> styleMap, HSSFWorkbook retVal)
        {
            // manage a list of merged zone in order to not insert two times a
            // merged zone
            List<CellRangeAddress> mergedRegions = new List<CellRangeAddress>();
            destRow.Height = srcRow.Height;
            // pour chaque row
            for (int j = srcRow.FirstCellNum; j <= srcRow.LastCellNum; j++)
            {
                XSSFCell oldCell = (XSSFCell)srcRow.GetCell(j); // ancienne cell
                HSSFCell newCell = (HSSFCell)destRow.GetCell(j); // new cell
                if (oldCell != null)
                {
                    if (newCell == null)
                    {
                        newCell = (HSSFCell)destRow.CreateCell(j);
                    }
                    // copy chaque cell
                    CopyCell(oldCell, newCell, styleMap, retVal);
                    // copy les informations de fusion entre les cellules
                    CellRangeAddress mergedRegion = GetMergedRegion(srcSheet, srcRow.RowNum,
                            (short)oldCell.ColumnIndex);

                    if (mergedRegion != null)
                    {
                        CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow,
                                mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);
                        if (IsNewMergedRegion(newMergedRegion, mergedRegions))
                        {
                            mergedRegions.Add(newMergedRegion);
                            destSheet.AddMergedRegion(newMergedRegion);
                        }

                        if (newMergedRegion.FirstColumn == 0 && newMergedRegion.LastColumn == 6 && newMergedRegion.FirstRow == newMergedRegion.LastRow)
                        {
                            HSSFCellStyle style2 = (HSSFCellStyle)retVal.CreateCellStyle();
                            style2.VerticalAlignment = VerticalAlignment.Center;
                            style2.Alignment = HorizontalAlignment.Left;
                            style2.FillForegroundColor = HSSFColor.Teal.Index;
                            style2.FillPattern = FillPattern.SolidForeground;

                            for (int i = destRow.FirstCellNum; i <= destRow.LastCellNum; i++)
                            {
                                if (destRow.GetCell(i) != null)
                                    destRow.GetCell(i).CellStyle = style2;
                            }
                        }
                    }
                }
            }
        }

        private static void CopyCell(XSSFCell oldCell, HSSFCell newCell, Dictionary<int, XSSFCellStyle> styleMap, HSSFWorkbook retVal)
        {
            if (styleMap != null)
            {
                int stHashCode = oldCell.CellStyle.Index;
                XSSFCellStyle sourceCellStyle = null;
                if (styleMap.TryGetValue(stHashCode, out sourceCellStyle)) { }

                HSSFCellStyle destnCellStyle = (HSSFCellStyle)newCell.CellStyle;
                if (sourceCellStyle == null)
                {
                    sourceCellStyle = (XSSFCellStyle)oldCell.Sheet.Workbook.CreateCellStyle();
                }
                // destnCellStyle.CloneStyleFrom(oldCell.CellStyle);
                if (!styleMap.Any(p => p.Key == stHashCode))
                {
                    styleMap.Add(stHashCode, sourceCellStyle);
                }

                destnCellStyle.VerticalAlignment = VerticalAlignment.Top;
                newCell.CellStyle = (HSSFCellStyle)destnCellStyle;
            }
            switch (oldCell.CellType)
            {
                case CellType.String:
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;

                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;

                case CellType.Blank:
                    newCell.SetCellType(CellType.Blank);
                    break;

                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;

                case CellType.Error:
                    newCell.SetCellErrorValue(FormulaError.ForInt(oldCell.ErrorCellValue));
                    break;

                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;

                default:
                    break;
            }
        }

        private static CellRangeAddress GetMergedRegion(XSSFSheet sheet, int rowNum, short cellNum)
        {
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress merged = sheet.GetMergedRegion(i);
                if (merged.IsInRange(rowNum, cellNum))
                {
                    return merged;
                }
            }
            return null;
        }

        /*
        private static bool IsNewMergedRegion(CellRangeAddress newMergedRegion, List<CellRangeAddress> mergedRegions)
        {
            return !mergedRegions.Contains(newMergedRegion);
        }
        */

        private static bool IsNewMergedRegion(CellRangeAddress newMergedRegion, List<CellRangeAddress> mergedRegions)
        {
            return !mergedRegions.Any(r => r.FirstColumn == newMergedRegion.FirstColumn && r.LastColumn == newMergedRegion.LastColumn && r.FirstRow == newMergedRegion.FirstRow && r.LastRow == newMergedRegion.LastRow);
        }
    }
}