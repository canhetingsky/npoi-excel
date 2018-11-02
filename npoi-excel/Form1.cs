﻿using System;
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
using System.Diagnostics;
using System.Threading;

namespace npoi_excel
{
    public partial class Form1 : Form
    {

        public struct Order_t
        {
            public string order_number;
            public string order_name;
            public string order_shipping_info;
            public int order_A_number;
            public int order_B_number;
            public int order_C_number;
            public int order_D_number;
            public int order_sum_number;

        };
        private Order_t order = new Order_t();  //提取Excel中信息的结构体
        private string folderName = null;   //要处理的目标文件夹
        private string[] fileNames = null;   //目标文件夹下的所有符合文件
        private int serialNumber = 0 ;      //序号，自增
        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            Thread Handle = new Thread(new ThreadStart(startProcessing));
            Handle.Start();
        }

        private void startProcessing()
        {
            fileNames = Directory.GetFiles(folderName, "*.xls", SearchOption.AllDirectories);
            Debug.Write("共有" + fileNames.Length + "个文件待处理");
            /*当使用星号通配符searchPattern如"*.txt"，指定的扩展中的字符数会影响搜索，如下所示：
            如果指定的扩展名为 3 个字符，该方法将返回具有指定扩展名以扩展名的文件。 例如，"*.xls"返回"book.xls"和"book.xlsx"。
            在所有其他情况下，该方法返回与指定的扩展完全匹配的文件。 例如，"*.ai"返回"file.ai"而不是"file.aif"。*/

            FileStream modeFile_Handover = new FileStream(@"template\实物ID交接单.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook workbook_Handover = new XSSFWorkbook(modeFile_Handover);
            modeFile_Handover.Close();
            XSSFSheet sheet_Handover = (XSSFSheet)workbook_Handover.GetSheetAt(0); //获取工作表 

            FileStream modeFile_Boxsign = new FileStream(@"template\实物ID箱签.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook workbook_Boxsign = new XSSFWorkbook(modeFile_Boxsign);
            modeFile_Boxsign.Close();

            foreach (string file in fileNames)
            {
                order = read_OrderFromExcel(file);
                serialNumber++;
                if (progressBar1.InvokeRequired)//不同线程为true，所以这里是true
                {
                    BeginInvoke(new Action<int>(target => { progressBar1.Value = target; }), serialNumber);
                }
                Debug.Write("@" + serialNumber + file + "\r\n");
                sheet_Handover = write_HandoverToExcel(sheet_Handover, order);
                write_BoxsignToExcel(workbook_Boxsign, order);
            }
            FileStream file_Handover = new FileStream(@"E:\excel-file\" + String.Format("{0:0000}", serialNumber) + "实物ID交接单.xlsx", FileMode.Create);
            workbook_Handover.Write(file_Handover);
            file_Handover.Close();
            workbook_Handover.Close();
        }

        private Order_t read_OrderFromExcel(string fileName)
        {
            
            Order_t file_order = new Order_t();
            IWorkbook workbook = null;  //新建IWorkbook对象
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook
                fileStream.Close();
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook
            }
            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表
            file_order.order_number = sheet.GetRow(1).GetCell(1).ToString();
            file_order.order_name = sheet.GetRow(1).GetCell(4).ToString();
            file_order.order_shipping_info = sheet.GetRow(3).GetCell(1).ToString();
            file_order.order_A_number = Convert.ToInt32(sheet.GetRow(5).GetCell(1).ToString());
            file_order.order_B_number = Convert.ToInt32(sheet.GetRow(6).GetCell(1).ToString());
            file_order.order_C_number = Convert.ToInt32(sheet.GetRow(7).GetCell(1).ToString());
            file_order.order_D_number = Convert.ToInt32(sheet.GetRow(8).GetCell(1).ToString());
            file_order.order_sum_number = Convert.ToInt32(sheet.GetRow(4).GetCell(1).ToString());

            Debug.Write("订单编号：" + file_order.order_number);
            Debug.Write("变电站或馈线名称：" + file_order.order_name);
            Debug.Write("发货信息：" + file_order.order_shipping_info);
            Debug.Write("标签A数量：" + file_order.order_A_number);
            Debug.Write("标签B数量：" + file_order.order_B_number);
            Debug.Write("标签C数量：" + file_order.order_C_number);
            Debug.Write("标签D数量：" + file_order.order_D_number);
            Debug.Write("标签总数量：" + file_order.order_sum_number);
            Debug.Write("\r\n");
            workbook.Close();
            return file_order;
        }

        private XSSFSheet write_HandoverToExcel(XSSFSheet sheet, Order_t order)
        {
            sheet.GetRow(2 + 4 * serialNumber).GetCell(0).SetCellValue(String.Format("{0:0000}", serialNumber));     //填写序号
            sheet.GetRow(2 + 4 * serialNumber).GetCell(1).SetCellValue(order.order_number);          //填写订单编号
            sheet.GetRow(2 + 4 * serialNumber).GetCell(2).SetCellValue("A型标签");                   //填写A型标签文本
            sheet.GetRow(2 + 4 * serialNumber).GetCell(4).SetCellValue(order.order_A_number);        //填写A型标签数量
            sheet.GetRow(2 + 4 * serialNumber + 1).GetCell(2).SetCellValue("B型标签");               //填写B型标签文本
            sheet.GetRow(2 + 4 * serialNumber + 1).GetCell(4).SetCellValue(order.order_B_number);    //填写B型标签数量
            sheet.GetRow(2 + 4 * serialNumber + 2).GetCell(2).SetCellValue("C型标签");               //填写C型标签文本
            sheet.GetRow(2 + 4 * serialNumber + 2).GetCell(4).SetCellValue(order.order_C_number);    //填写C型标签数量
            sheet.GetRow(2 + 4 * serialNumber + 3).GetCell(2).SetCellValue("D型标签");              //填写D型标签文本
            sheet.GetRow(2 + 4 * serialNumber + 3).GetCell(4).SetCellValue(order.order_D_number);              //填写D型标签数量
            sheet.GetRow(2 + 4 * serialNumber).GetCell(3).SetCellValue("1/1");     //填写箱号    
            return sheet;
        }

        private void write_BoxsignToExcel(XSSFWorkbook workbook, Order_t order)
        {
            XSSFSheet modeSheet_Boxsign = (XSSFSheet)workbook.GetSheetAt(0); //获取工作表
            modeSheet_Boxsign.GetRow(0).GetCell(1).SetCellValue(String.Format("{0:0000}", serialNumber));     //填写序号
            modeSheet_Boxsign.GetRow(1).GetCell(1).SetCellValue(order.order_number);     //填写订单编号
            modeSheet_Boxsign.GetRow(0).GetCell(4).SetCellValue(order.order_name);       //填写变电站或馈线名称

            string[] str = { " " };
            string[] string_split_word = order.order_shipping_info.Split(str, StringSplitOptions.RemoveEmptyEntries);
            modeSheet_Boxsign.GetRow(2).GetCell(1).SetCellValue(string_split_word[0]);   //填写收货单位
            modeSheet_Boxsign.GetRow(5).GetCell(1).SetCellValue(string_split_word[1]);   //填写联系人
            modeSheet_Boxsign.GetRow(5).GetCell(4).SetCellValue(string_split_word[2]);   //填写电话
            modeSheet_Boxsign.GetRow(4).GetCell(1).SetCellValue(string_split_word[3]);   //填写发货地址

            modeSheet_Boxsign.GetRow(7).GetCell(1).SetCellValue(order.order_A_number);       //填写A型标签数量
            modeSheet_Boxsign.GetRow(8).GetCell(1).SetCellValue(order.order_B_number);       //填写B型标签数量
            modeSheet_Boxsign.GetRow(9).GetCell(1).SetCellValue(order.order_C_number);       //填写C型标签数量
            modeSheet_Boxsign.GetRow(10).GetCell(1).SetCellValue(order.order_D_number);      //填写D型标签数量
            modeSheet_Boxsign.GetRow(10).GetCell(4).SetCellValue(order.order_sum_number);   //填写总数量

            
            FileStream file_Boxsign = new FileStream(@"E:\excel-file\实物ID箱签\" + String.Format("{0:0000}", serialNumber) + "-" + order.order_number + "-实物ID箱签.xlsx", FileMode.Create);
            workbook.Write(file_Boxsign);
            file_Boxsign.Close();
            workbook.Close();
        }

        private void labelFiles_Click(object sender, EventArgs e)
        {
            //选择要处理的文件
            //OpenFileDialog loadFiles = new OpenFileDialog();
            //loadFiles.Filter = "Excel文件|*.xls;*.xlsx";
            //loadFiles.Multiselect = false;
            //loadFiles.Title = "选择Excel文件";
            //if (loadFiles.ShowDialog() == DialogResult.OK)
            //{
            //    fileName = loadFiles.FileName;
            //    labelFiles.Text = fileName;
            //    Debug.Write(fileName);
            //}

            //选择要处理的文件夹
            FolderBrowserDialog loadFolder = new FolderBrowserDialog();
            loadFolder.Description = "请选择文件夹路径";
            loadFolder.ShowNewFolderButton = false;
            if (loadFolder.ShowDialog() == DialogResult.OK)
            {
                folderName = loadFolder.SelectedPath;
                labelFiles.Text = folderName;
                progressBar1.Maximum = folderName.Length;
            }
        }
    }
}
