using ExcelParser.Constants;
using ExcelParser.Model;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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

            var list = ToList(package);
            return CommonUtil.MappingDatatoDataGridViewForList(list);
        }

        public static void ExportExcel(DataTable dt, FileInfo fileInfo)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add(CommonConstants.WorkSheetName);
                workSheet.Cells.LoadFromDataTable(dt, true);
                var firstRow = workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column];
                firstRow.Style.Fill.SetBackground(Color.Orange);
                excelPackage.Save();
            }
        }

        public static List<OpenCardModel> ToList(ExcelPackage package)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets.First();
            List<OpenCardModel> list = new();
            OpenCardModel openCardModel;

            for (int rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
            {
                var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];

                //Eğer tüm satırlar boşsa döngüden çık
                if (row.Any(c => c.Value == null))
                {
                    break;
                }

                openCardModel = new OpenCardModel
                {
                    model = ((object[,])row.Value)[0, 3] != null ? ((object[,])row.Value)[0, 3].ToString() : "",
                    option_name = ((object[,])row.Value)[0, 4] != null ? ((object[,])row.Value)[0, 4].ToString() : "",
                    option_value = ((object[,])row.Value)[0, 6] != null ? ((object[,])row.Value)[0, 6].ToString() : "",
                    manufacturer = ((object[,])row.Value)[0, 7] != null ? ((object[,])row.Value)[0, 7].ToString() : "",
                    category = ((object[,])row.Value)[0, 8] != null ? ((object[,])row.Value)[0, 8].ToString() : "",
                    name = ((object[,])row.Value)[0, 10] != null ? ((object[,])row.Value)[0, 10].ToString() : "",
                    description = ((object[,])row.Value)[0, 11] != null ? ((object[,])row.Value)[0, 11].ToString() : "",
                    price = ((object[,])row.Value)[0, 12] != null ? ((object[,])row.Value)[0, 12].ToString() : "",
                    quantity = ((object[,])row.Value)[0, 16] != null ? ((object[,])row.Value)[0, 16].ToString() : "0",
                    tax_class = ((object[,])row.Value)[0, 20] != null ? ((object[,])row.Value)[0, 20].ToString() : "",
                    weight = ((object[,])row.Value)[0, 21] != null ? ((object[,])row.Value)[0, 21].ToString() : "",
                    image = CommonUtil.ImageMapping(((object[,])row.Value)[0, 22] != null ? ((object[,])row.Value)[0, 22].ToString() : "", false),
                    additional_image = CommonUtil.ImageMapping(((object[,])row.Value)[0, 22] != null ? ((object[,])row.Value)[0, 22].ToString() : "", true),
                    option_type = CommonConstants.OptionType,
                    option_required = CommonConstants.OptionRequired,
                    option_subtract = CommonConstants.OptionSubtract,
                    option_quantity = ((object[,])row.Value)[0, 16] != null ? ((object[,])row.Value)[0, 16].ToString() : "0",
                    status = CommonConstants.Status

                };
                list.Add(openCardModel);

            }
            return list;
        }
    }
}
