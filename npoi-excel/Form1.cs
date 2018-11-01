using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace npoi_excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            HSSFWorkbook workbook2003 = new HSSFWorkbook(); //新建xls工作簿
            workbook2003.CreateSheet("Sheet1");  //新建3个Sheet工作表
            workbook2003.CreateSheet("Sheet2");
            workbook2003.CreateSheet("Sheet3");
            FileStream file2003 = new FileStream(@"E:\Excel2003.xls", FileMode.Create);
            workbook2003.Write(file2003);
            file2003.Close();  //关闭文件流
            workbook2003.Close();

            XSSFWorkbook workbook2007 = new XSSFWorkbook();  //新建xlsx工作簿
            workbook2007.CreateSheet("Sheet1");
            workbook2007.CreateSheet("Sheet2");
            workbook2007.CreateSheet("Sheet3");
            FileStream file2007 = new FileStream(@"E:\Excel2007.xlsx", FileMode.Create);
            workbook2007.Write(file2007);
            file2007.Close();
            workbook2007.Close();
        }
    }
}
