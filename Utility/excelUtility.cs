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
        public static byte[] ExportDataTableToExcel(DataTable dataTable)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                }

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int column = 0; column < dataTable.Columns.Count; column++)
                    {
                        worksheet.Cells[row + 2, column + 1].Value = dataTable.Rows[row][column];
                    }
                }
                return excelPackage.GetAsByteArray();
            }
        }
    }
}
