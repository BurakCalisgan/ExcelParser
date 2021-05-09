using ExcelParser.Constants;
using ExcelParser.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Utility
{
    public static class CommonUtil
    {
        /// <summary>
        /// Quantity calculate and list to datatable mapping
        /// </summary>
        /// <param name="openCardList"></param>
        /// <returns></returns>
        public static DataTable MappingDatatoDataGridViewForList(List<OpenCardModel> openCardList)
        {
            List<OpenCardModel> result = openCardList
                                        .GroupBy(l => l.model)
                                        .Select(cl => new OpenCardModel
                                        {
                                            model = cl.First().model,
                                            quantity = cl.Sum(c => Convert.ToInt32(c.quantity)).ToString(),
                                        }).ToList();

            for (int i = 0; i < result.Count; i++)
            {
                openCardList.Where(x => x.model == result[i].model).FirstOrDefault().quantity = result[i].quantity;
            }

            DataTable dt = ToDataTable(openCardList);
            return dt;
        }

        /// <summary>
        /// List to DataTable Converter Function
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="items"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        /// <summary>
        /// Image Mapping
        /// </summary>
        /// <param name="allImage"></param>
        /// <param name="isAdditionalImage"></param>
        /// <returns></returns>
        public static string ImageMapping(string allImage, bool isAdditionalImage)
        {
            string newValue;
            if (!string.IsNullOrEmpty(allImage))
            {
                if (allImage.Contains(";"))
                {
                    newValue = allImage.Replace(";", ":::");
                    if (isAdditionalImage)
                    {
                        return newValue;
                    }
                    else
                    {
                        return newValue.Split(":::").FirstOrDefault();
                    }
                }
            }
            return "";
        }

        /// <summary>
        /// Columns Adding For OpenCard
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        //public static DataTable AddingHeaderColumnsOpenCardCsvFormat(DataTable dt)
        //{
        //    //OpenCard Columns Add
        //    dt.Columns.Add(OpenCardConstants.model);
        //    dt.Columns.Add(OpenCardConstants.sku);
        //    dt.Columns.Add(OpenCardConstants.upc);
        //    dt.Columns.Add(OpenCardConstants.ean);
        //    dt.Columns.Add(OpenCardConstants.jan);
        //    dt.Columns.Add(OpenCardConstants.isbn);
        //    dt.Columns.Add(OpenCardConstants.mpn);
        //    dt.Columns.Add(OpenCardConstants.name);
        //    dt.Columns.Add(OpenCardConstants.description);
        //    dt.Columns.Add(OpenCardConstants.category);
        //    dt.Columns.Add(OpenCardConstants.image);
        //    dt.Columns.Add(OpenCardConstants.additional_image);
        //    dt.Columns.Add(OpenCardConstants.manufacturer);
        //    dt.Columns.Add(OpenCardConstants.price);
        //    dt.Columns.Add(OpenCardConstants.tax_class);
        //    dt.Columns.Add(OpenCardConstants.quantity);
        //    dt.Columns.Add(OpenCardConstants.minimum);
        //    dt.Columns.Add(OpenCardConstants.subtract);
        //    dt.Columns.Add(OpenCardConstants.stock_status);
        //    dt.Columns.Add(OpenCardConstants.status);
        //    dt.Columns.Add(OpenCardConstants.date_available);
        //    dt.Columns.Add(OpenCardConstants.shipping);
        //    dt.Columns.Add(OpenCardConstants.weight);
        //    dt.Columns.Add(OpenCardConstants.length);
        //    dt.Columns.Add(OpenCardConstants.width);
        //    dt.Columns.Add(OpenCardConstants.height);
        //    dt.Columns.Add(OpenCardConstants.meta_keyword);
        //    dt.Columns.Add(OpenCardConstants.meta_title);
        //    dt.Columns.Add(OpenCardConstants.meta_description);
        //    dt.Columns.Add(OpenCardConstants.sort_order);
        //    dt.Columns.Add(OpenCardConstants.tag);
        //    dt.Columns.Add(OpenCardConstants.product_url);
        //    dt.Columns.Add(OpenCardConstants.points);
        //    dt.Columns.Add(OpenCardConstants.related_product);
        //    dt.Columns.Add(OpenCardConstants.layout);
        //    dt.Columns.Add(OpenCardConstants.location);
        //    dt.Columns.Add(OpenCardConstants.date_added);
        //    dt.Columns.Add(OpenCardConstants.date_modified);
        //    dt.Columns.Add(OpenCardConstants.feed_product_id);
        //    dt.Columns.Add(OpenCardConstants.import_id);
        //    dt.Columns.Add(OpenCardConstants.import_active_product);
        //    dt.Columns.Add(OpenCardConstants.currency_id);
        //    dt.Columns.Add(OpenCardConstants.skip_import);
        //    dt.Columns.Add(OpenCardConstants.meta_robots);
        //    dt.Columns.Add(OpenCardConstants.seo_keyword);
        //    dt.Columns.Add(OpenCardConstants.seo_h1);
        //    dt.Columns.Add(OpenCardConstants.seo_h2);
        //    dt.Columns.Add(OpenCardConstants.seo_h3);
        //    dt.Columns.Add(OpenCardConstants.image_title);
        //    dt.Columns.Add(OpenCardConstants.image_alt);
        //    dt.Columns.Add(OpenCardConstants.option_name);
        //    dt.Columns.Add(OpenCardConstants.option_type);
        //    dt.Columns.Add(OpenCardConstants.option_value);
        //    dt.Columns.Add(OpenCardConstants.option_required);
        //    dt.Columns.Add(OpenCardConstants.option_image);
        //    dt.Columns.Add(OpenCardConstants.option_sort_order);
        //    dt.Columns.Add(OpenCardConstants.option_quantity);
        //    dt.Columns.Add(OpenCardConstants.option_subtract);
        //    dt.Columns.Add(OpenCardConstants.option_price);
        //    dt.Columns.Add(OpenCardConstants.option_points);
        //    dt.Columns.Add(OpenCardConstants.option_weight);

        //    return dt;

        //}

        /// <summary>
        /// Manuel DataTableMapping
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="ws"></param>
        /// <returns></returns>
        //public static DataTable MappingDatatoDataGridView(DataTable dt, ExcelWorksheet ws)
        //{
        //    //var productList = ExcelPackageUtil.ToList();
        //    var firstRow = ws.Cells[1, 1, 1, ws.Dimension.End.Column];
        //    int columnId;
        //    string columnName;

        //    //TO DO : Data mapping yapılacak. Resim business'ı yapılacak. Quantity hesaplanacak. 
        //    for (int rowNumber = 2; rowNumber < ws.Dimension.End.Row; rowNumber++)
        //    {
        //        var row = ws.Cells[rowNumber, 1, rowNumber, ws.Dimension.End.Column];
                
        //        var newRow = dt.NewRow();

        //        for (int i = 0; i < row.Count(); i++)
        //        {
        //            columnId = ws.Cells[1, ws.Dimension.Start.Column, 1, ws.Dimension.End.Column].First(c => c.Value.ToString() == ((object[,])firstRow.Value)[0, i].ToString()).Start.Column;
        //            columnName = GetOpenCardColumnNameByTrendyolColumnId(columnId);
        //            if (columnName != "no_column")
        //            {
        //                if (columnName == OpenCardConstants.additional_image)   //main image and additional image seperation
        //                {
        //                    newRow[columnName] = ImageMapping(((object[,])row.Value)[0, i].ToString(), true);
        //                    newRow[OpenCardConstants.image] = ImageMapping(((object[,])row.Value)[0, i].ToString(),false);
        //                }
        //                else
        //                {
        //                    newRow[columnName] = ((object[,])row.Value)[0, i].ToString();
        //                }
                        
        //            }

        //            newRow[OpenCardConstants.option_type] = CommonConstants.OptionType;
        //            newRow[OpenCardConstants.option_required] = CommonConstants.OptionRequired;
        //            newRow[OpenCardConstants.option_subtract] = CommonConstants.OptionSubtract;
        //            newRow[OpenCardConstants.status] = CommonConstants.Status;
        //        }
        //        dt.Rows.Add(newRow);
        //    }

        //    return dt;
        //}

        //public static string GetOpenCardColumnNameByTrendyolColumnId(int columnId)
        //{
        //    switch (columnId)
        //    {
        //        case 4:
        //            return OpenCardConstants.model;
        //        case 5:
        //            return OpenCardConstants.option_name;
        //        case 6:
        //            return OpenCardConstants.option_value;
        //        case 8:
        //            return OpenCardConstants.manufacturer;
        //        case 9:
        //            return OpenCardConstants.category;
        //        case 11:
        //            return OpenCardConstants.name;
        //        case 12:
        //            return OpenCardConstants.description;
        //        case 13:
        //            return OpenCardConstants.price;
        //        case 17:
        //            return OpenCardConstants.quantity;
        //        case 21:
        //            return OpenCardConstants.tax_class;
        //        case 22:
        //            return OpenCardConstants.weight;
        //        case 23:
        //            return OpenCardConstants.additional_image;
        //        case 25:
        //            return OpenCardConstants.status;

        //        default:
        //            return "no_column";
        //    }
        //}

       
    }
}
