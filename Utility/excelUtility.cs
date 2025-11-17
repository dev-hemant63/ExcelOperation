using Microsoft.AspNetCore.Http;
using System.ComponentModel;
using System.Data;
using System;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace ExcelOperation.Utility
{
    public static class excelUtility
    {
        public static DataTable ExcelDataToDataTable(IFormFile file, string sheetName, bool hasHeader = true)
        {
            var dt = new DataTable();

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var stream = file.OpenReadStream())
                using (var xlPackage = new ExcelPackage(stream))
                {
                    var worksheet = xlPackage.Workbook.Worksheets[sheetName];

                    dt = worksheet.Cells[1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].ToDataTable(c =>
                    {
                        c.FirstRowIsColumnNames = hasHeader;
                    });
                }
            }
            catch (Exception ex)
            {

            }

            return dt;
        }
 public static byte[] ExportDataTableToExcel(DataTable dataTable, string title)
{
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    using (ExcelPackage excelPackage = new ExcelPackage())
    {
        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

        int totalColumns = dataTable.Columns.Count;

        // ======================================
        // 1. TITLE (ROW 1)
        // ======================================
        worksheet.Cells[1, 1].Value = title;
        worksheet.Cells[1, 1, 1, totalColumns].Merge = true;

        worksheet.Cells[1, 1].Style.Font.Bold = true;
        worksheet.Cells[1, 1].Style.Font.Size = 16;
        worksheet.Cells[1, 1].Style.Font.Color.SetColor(Color.DarkBlue);
        worksheet.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
        worksheet.Cells[1, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        worksheet.Cells[1, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

        worksheet.Row(1).Height = 28;

        // ======================================
        // 2. HEADER (ROW 2)
        // ======================================
        for (int i = 0; i < totalColumns; i++)
        {
            worksheet.Cells[2, i + 1].Value = dataTable.Columns[i].ColumnName;
            worksheet.Cells[2, i + 1].Style.Font.Bold = true;

            worksheet.Cells[2, i + 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            worksheet.Cells[2, i + 1].Style.Fill.BackgroundColor.SetColor(Color.SteelBlue);
            worksheet.Cells[2, i + 1].Style.Font.Color.SetColor(Color.White);
            worksheet.Cells[2, i + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        // ======================================
        // 3. DATA (START ROW 3)
        // ======================================
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < totalColumns; col++)
            {
                worksheet.Cells[row + 3, col + 1].Value = dataTable.Rows[row][col];
            }
        }

        // ======================================
        // 4. APPLY ALL BORDERS
        // ======================================
        int totalRows = dataTable.Rows.Count + 2; // title = row1, header = row2

        using (var range = worksheet.Cells[1, 1, totalRows, totalColumns])
        {
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            range.Style.Border.Top.Color.SetColor(Color.Black);
            range.Style.Border.Bottom.Color.SetColor(Color.Black);
            range.Style.Border.Left.Color.SetColor(Color.Black);
            range.Style.Border.Right.Color.SetColor(Color.Black);
        }

        // Auto-fit columns
        worksheet.Cells.AutoFitColumns();

        return excelPackage.GetAsByteArray();
    }
}

        }
    }
}
