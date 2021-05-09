using ExcelParser.Constants;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelParser.Utility
{
    public static class ExcelPackageUtil
    {
        public static DataTable ReadExcel(OpenFileDialog file)
        {
            FileInfo[] files = file.FileNames.Select(f => new FileInfo(f)).ToArray();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(files[0]);
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            DataTable dt = new DataTable();
            foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
            {
                //TO DO : Buraya column mapping işlemleri gelecek.
                dt.Columns.Add(firstRowCell.Text);
            }

            //TO DO : Data mapping yapılacak. Resim business'ı yapılacak. Quantity hesaplanacak. 
            for (int rowNumber = 2; rowNumber < workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                var newRow = dt.NewRow();
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }
                dt.Rows.Add(newRow);
            }
            return dt;
        }

        public static void ExportExcel(DataTable dt, FileInfo fileInfo)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(CommonConstants.WorkSheetName);
                worksheet.Cells.LoadFromDataTable(dt, true);
                excelPackage.Save();
            }
        }
    }
}
