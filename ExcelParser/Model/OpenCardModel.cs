using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.Model
{
    public class OpenCardModel
    {
        public string model { get; set; }
        public string sku { get; set; }
        public string upc { get; set; }
        public string ean { get; set; }
        public string jan { get; set; }
        public string isbn { get; set; }
        public string mpn { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string category { get; set; }
        public string image { get; set; }
        public string additional_image { get; set; }
        public string manufacturer { get; set; }
        public string price { get; set; }
        public string tax_class { get; set; }
        public string quantity { get; set; }
        public string minimum { get; set; }
        public string subtract { get; set; }
        public string stock_status { get; set; }
        public string status { get; set; }
        public string date_available { get; set; }
        public string shipping { get; set; }
        public string weight { get; set; }
        public string length { get; set; }
        public string width { get; set; }
        public string height { get; set; }
        public string meta_keyword { get; set; }
        public string meta_title { get; set; }
        public string meta_description { get; set; }
        public string sort_order { get; set; }
        //public string seo_keyword = "seo_keyword";  2 kere yazilmis hata olabilir.
        public string tag { get; set; }
        public string product_url { get; set; }
        public string points { get; set; }
        public string related_product { get; set; }
        public string layout { get; set; }
        public string location { get; set; }
        public string date_added { get; set; }
        public string date_modified { get; set; }
        public string feed_product_id { get; set; }
        public string import_id { get; set; }
        public string import_active_product { get; set; }
        public string currency_id { get; set; }
        public string skip_import { get; set; }
        public string meta_robots { get; set; }
        public string seo_keyword { get; set; }
        public string seo_h1 { get; set; }
        public string seo_h2 { get; set; }
        public string seo_h3 { get; set; }
        public string image_title { get; set; }
        public string image_alt { get; set; }
        public string option_name { get; set; }
        public string option_type { get; set; }
        public string option_value { get; set; }
        public string option_required { get; set; }
        public string option_image { get; set; }
        public string option_sort_order { get; set; }
        public string option_quantity { get; set; }
        public string option_subtract { get; set; }
        public string option_price { get; set; }
        public string option_points { get; set; }
        public string option_weight { get; set; }
    }
}
