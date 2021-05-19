using ExcelParser.Constants;
using ExcelParser.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelParser
{
    public partial class Home : Form
    {
        public Home()
        {
            InitializeComponent();
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
            if (file.ShowDialog() == DialogResult.OK) //if there is a file choosen by the user  
            {
                filePath = file.FileName; //get the path of the file  
                fileExt = Path.GetExtension(filePath); //get the file extension  

                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new();
                        dtExcel = ExcelPackageUtil.ReadExcel(file); //read excel file  
                        dtgExcelData.Visible = true;
                        dtgExcelData.DataSource = dtExcel;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = CommonConstants.FileTypes })
                {
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        //var fileInfo = new FileInfo(sfd.FileName);
                        //ExcelPackageUtil.ExportExcel((DataTable)dtgExcelData.DataSource, fileInfo);
                        var dt = (DataTable)dtgExcelData.DataSource;
                        dt.ToCSV(sfd.FileName);
                        MessageBox.Show("Operation has been succesfully completed.");
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Operation has not been succesfully completed. --> " + ex.Message);
            }

        }
    }
}
