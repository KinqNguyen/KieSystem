using KieSystem.DTOs;
using KieSystem.Interface;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;

namespace KieSystem.Services
{
    public class ExcelService : IExcelService
    {

        public byte[] RawExportExcel(IEnumerable<string> columns, string HeaderTitle, IEnumerable<BlogDTO> data)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            IRow titleRow = sheet.CreateRow(0);
            ICell titleCell = titleRow.CreateCell(0);
            titleCell.SetCellValue(HeaderTitle);


            // Create header row
            IRow headerRow = sheet.CreateRow(1);
            for (int i = 0; i < columns.Count(); i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(columns.ElementAt(i));
            }

            // Create data rows
            int rowIndex = 2;
            foreach (var item in data)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int i = 0; i < columns.Count(); i++)
                {
                    string column = columns.ElementAt(i);
                    ICell cell = row.CreateCell(i);

                    var property = typeof(BlogDTO).GetProperties()
                        .FirstOrDefault(p => p.Name.Equals(column, StringComparison.OrdinalIgnoreCase));
                    if (property != null)
                    {
                        var value = property.GetValue(item)?.ToString() ?? string.Empty;
                        cell.SetCellValue(value);
                    }
                }
            }

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }

        public byte[] ExportExcel(ExportExcelDTO exportDto, IEnumerable<BlogDTO> data)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Header Title
            if (exportDto.generalInfo.HeaderTitle.Merge.Enable)
            {
                IRow titleRow = sheet.CreateRow(0);
                ICell titleCell = titleRow.CreateCell(0);
                titleCell.SetCellValue(exportDto.generalInfo.HeaderTitle.HeaderName);

                var mergeRegion = new CellRangeAddress(0, 0, 0, exportDto.generalInfo.HeaderTitle.Merge.Count - 1);
                sheet.AddMergedRegion(mergeRegion);

                ICellStyle headerStyle = workbook.CreateCellStyle();
                IFont headerFont = workbook.CreateFont();
                headerFont.IsBold = exportDto.generalInfo.HeaderTitle.Font.Bold;
                headerFont.IsItalic = exportDto.generalInfo.HeaderTitle.Font.Italic;
                headerFont.Underline = exportDto.generalInfo.HeaderTitle.Font.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
                headerStyle.SetFont(headerFont);
                headerStyle.FillForegroundColor = FindClosestNPOIColor(exportDto.generalInfo.HeaderTitle.Color);
                headerStyle.FillPattern = FillPattern.SolidForeground;
                headerStyle.BorderBottom = BorderStyle.Thick;
                headerStyle.BorderTop = BorderStyle.Thick;
                headerStyle.BorderLeft = BorderStyle.Thick;
                headerStyle.BorderRight = BorderStyle.Thick;

                titleCell.CellStyle = headerStyle;
            }

            IRow headerRow = sheet.CreateRow(exportDto.generalInfo.HeaderTitle.Merge.Enable ? 1 : 0);

            // Determine the columns to export
            var columnsToExport = new List<string>();


                columnsToExport.AddRange(exportDto.generalInfo.ColumnExport.Where(c => c != "BonusTotal"));

            // Ensure UserLevel and TotalView are included if BonusTotal is requested
            if (exportDto.generalInfo.ColumnExport.Contains("BonusTotal"))
            {
                if (!exportDto.generalInfo.ColumnExport.Contains("Userlevel"))
                {
                    columnsToExport.Add("UserLevel");
                }
                if (!exportDto.generalInfo.ColumnExport.Contains("Totalview"))
                {
                    columnsToExport.Add("TotalView");
                }

                // Add the rest of the columns and ensure BonusTotal is last
                columnsToExport.Add("BonusTotal");
            }
            // Add columns to header row
            foreach (var column in columnsToExport)
            {
                int index = columnsToExport.IndexOf(column);
                ICell cell = headerRow.CreateCell(index);
                cell.SetCellValue(column);

                ICellStyle cellStyle = workbook.CreateCellStyle();
                IFont font = workbook.CreateFont();
                font.IsBold = exportDto.generalInfo.HeaderTitle.Font.Bold;
                font.IsItalic = exportDto.generalInfo.HeaderTitle.Font.Italic;
                font.Underline = exportDto.generalInfo.HeaderTitle.Font.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
                cellStyle.SetFont(font);
                cellStyle.BorderBottom = BorderStyle.Thick;
                cellStyle.BorderTop = BorderStyle.Thick;
                cellStyle.BorderLeft = BorderStyle.Thick;
                cellStyle.BorderRight = BorderStyle.Thick;

                cell.CellStyle = cellStyle;
                AutoSizeColumn(sheet, index); // tương đương Alt + H O I
            }

            int rowIndex = exportDto.generalInfo.HeaderTitle.Merge.Enable ? 2 : 1;
            foreach (var item in data)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                foreach (var column in columnsToExport)
                {
                    int columnIndex = columnsToExport.IndexOf(column);
                    ICell cell = row.CreateCell(columnIndex);

                    if (column.Equals("BonusTotal", StringComparison.OrdinalIgnoreCase))
                    {
                        // Add formula for BonusTotal
                        string userLevelColumnLetter = GetExcelColumnLetter(columnsToExport.IndexOf("UserLevel"));
                        string totalViewColumnLetter = GetExcelColumnLetter(columnsToExport.IndexOf("TotalView"));
                        string formula = $"{userLevelColumnLetter}{rowIndex}*0.1*{totalViewColumnLetter}{rowIndex}*1000";
                        cell.CellFormula = formula;
                    }
                    else
                    {
                        var property = typeof(BlogDTO).GetProperties()
                            .FirstOrDefault(p => p.Name.Equals(column, StringComparison.OrdinalIgnoreCase));
                        if (property != null)
                        {
                            var value = property.GetValue(item)?.ToString() ?? string.Empty;
                            cell.SetCellValue(value);
                        }
                    }

                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.BorderBottom = BorderStyle.Thick;
                    cellStyle.BorderTop = BorderStyle.Thick;
                    cellStyle.BorderLeft = BorderStyle.Thick;
                    cellStyle.BorderRight = BorderStyle.Thick;
                    cell.CellStyle = cellStyle;
                }

                if (exportDto.generalInfo.Alternate.Enable && rowIndex % 2 == 0)
                {
                    ICellStyle cellStyle = workbook.CreateCellStyle();
                    cellStyle.FillForegroundColor = FindClosestNPOIColor(exportDto.generalInfo.Alternate.Color);
                    cellStyle.FillPattern = FillPattern.SolidForeground;
                    cellStyle.BorderBottom = BorderStyle.Thick;
                    cellStyle.BorderTop = BorderStyle.Thick;
                    cellStyle.BorderLeft = BorderStyle.Thick;
                    cellStyle.BorderRight = BorderStyle.Thick;

                    foreach (ICell cell in row.Cells)
                    {
                        cell.CellStyle = cellStyle;
                    }
                }
            }

            // Set column width for hidden or minimized columns
            if (!exportDto.generalInfo.ColumnExport.Contains("UserLevel"))
            {
                int userLevelIndex = columnsToExport.IndexOf("UserLevel");
                if (userLevelIndex >= 0)
                {
                    sheet.SetColumnWidth(userLevelIndex, 1 * 256); // Very narrow column
                }
            }

            if (!exportDto.generalInfo.ColumnExport.Contains("TotalView"))
            {
                int totalViewIndex = columnsToExport.IndexOf("TotalView");
                if (totalViewIndex >= 0)
                {
                    sheet.SetColumnWidth(totalViewIndex, 1 * 256); // Very narrow column
                }
            }

            // Set column width for other columns
            foreach (var column in exportDto.generalInfo.ColumnExport)
            {
                int index = columnsToExport.IndexOf(column);
                if (index >= 0)
                {
                    sheet.SetColumnWidth(index, 40 * 256);
                }
            }

            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }



        private string GetExcelColumnLetter(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }



        static short FindClosestNPOIColor(string hexColor)
        {
            // Chuyển đổi màu hex thành Color
            Color targetColor = ColorTranslator.FromHtml(hexColor);

            // Danh sách các màu được hỗ trợ bởi NPOI
            var npoiColors = new Dictionary<short, Color>
        {
            { IndexedColors.Black.Index, Color.Black },
            { IndexedColors.White.Index, Color.White },
            { IndexedColors.Red.Index, Color.Red },
            { IndexedColors.BrightGreen.Index, Color.Lime },
            { IndexedColors.Blue.Index, Color.Blue },
            { IndexedColors.Yellow.Index, Color.Yellow },
            { IndexedColors.Pink.Index, Color.Pink },
            { IndexedColors.Turquoise.Index, Color.Cyan },
            { IndexedColors.DarkRed.Index, Color.DarkRed },
            { IndexedColors.Green.Index, Color.Green },
            { IndexedColors.DarkBlue.Index, Color.DarkBlue },
            { IndexedColors.DarkYellow.Index, Color.Olive },
            { IndexedColors.Violet.Index, Color.Violet },
            { IndexedColors.Teal.Index, Color.Teal },
            { IndexedColors.Grey25Percent.Index, Color.FromArgb(64, 64, 64) },
            { IndexedColors.Grey50Percent.Index, Color.FromArgb(128, 128, 128) },
            { IndexedColors.CornflowerBlue.Index, Color.CornflowerBlue },
            { IndexedColors.Maroon.Index, Color.Maroon },
            { IndexedColors.LemonChiffon.Index, Color.LemonChiffon },
            { IndexedColors.Orchid.Index, Color.Orchid },
            { IndexedColors.Coral.Index, Color.Coral },
            { IndexedColors.RoyalBlue.Index, Color.RoyalBlue },
            { IndexedColors.LightCornflowerBlue.Index, Color.LightSteelBlue }
        };

            // Tìm màu gần nhất
            short closestIndex = npoiColors.Keys.First();
            double minDistance = double.MaxValue;

            foreach (var kvp in npoiColors)
            {
                double distance = GetColorDistance(targetColor, kvp.Value);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestIndex = kvp.Key;
                }
            }

            return closestIndex;
        }

        static double GetColorDistance(Color color1, Color color2)
        {
            int rDiff = color1.R - color2.R;
            int gDiff = color1.G - color2.G;
            int bDiff = color1.B - color2.B;
            return Math.Sqrt(rDiff * rDiff + gDiff * gDiff + bDiff * bDiff);
        }


    private void AutoSizeColumn(ISheet sheet, int columnIndex)
        {
            int columnWidth = 0;
            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    ICell cell = row.GetCell(columnIndex);
                    if (cell != null)
                    {
                        int length = cell.ToString().Length;
                        if (length > columnWidth)
                        {
                            columnWidth = length;
                        }
                    }
                }
            }
            // Set the width of the column in units of 1/256th of a character width
            sheet.SetColumnWidth(columnIndex, (columnWidth + 2) * 256); // Adding 2 for padding
        }
    }
}
