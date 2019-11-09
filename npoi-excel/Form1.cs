using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Collections;
using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.Collections.Generic;
//using Spire.Xls;

namespace npoi_excel
{
    public partial class Form1 : Form
    {

        public struct Order_t
        {
            public string order_number;
            public string order_name;
            public string order_shipping_info;
            public string order_A_number;
            public string order_B_number;
            public string order_C_number;
            public string order_D_number;
            public string order_sum_number;

        };
        private Order_t order = new Order_t();  //提取Excel中信息的结构体
        private string folderPath = null;   //要处理的目标文件夹
        private string[] fileNames = null;   //目标文件夹下的所有符合文件
        private int serialNumber = 0 ;      //序号，自增
        private int initialSerialNumber = 1;    //初始序号，默认为1
        new Thread Handle;   //处理excel的线程
        private string suffix =null;  //生成文件的后缀名

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            btnPackingList.Enabled = false;
            btnPrintBoxsign.Enabled = false;
            btnPrintHandover.Enabled = false;
            txbSerialNumber.Text = initialSerialNumber.ToString();

            comboBox1.SelectedIndex = 0;
        }
        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                initialSerialNumber = Convert.ToInt32(txbSerialNumber.Text);
                serialNumber = initialSerialNumber - 1;
            }
            catch (Exception)
            {
                MessageBox.Show("请输入正确的格式", "起始序号输入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            fileNames = Directory.GetFiles(folderPath, "*.xls", SearchOption.AllDirectories);   //得到目标文件夹下的所有Excel文件（包括子文件夹）
            Array.Sort(fileNames, new FileNameSort());  //对文件夹中的内容进行大小排序

            /*当使用星号通配符searchPattern如"*.txt"，指定的扩展中的字符数会影响搜索，如下所示：
            如果指定的扩展名为 3 个字符，该方法将返回具有指定扩展名以扩展名的文件。 例如，"*.xls"返回"book.xls"和"book.xlsx"。
            在所有其他情况下，该方法返回与指定的扩展完全匹配的文件。 例如，"*.ai"返回"file.ai"而不是"file.aif"。*/

            skinProgressBar1.Maximum = fileNames.Length;

            string[] subPath = new string[2];
            subPath[0] = folderPath + "/实物ID箱签/";
            subPath[1] = folderPath + "/实物ID箱签/";
            foreach (string item in subPath)
            {
                if (false == Directory.Exists(item))
                {
                    //创建实物ID箱签文件夹
                    Directory.CreateDirectory(item);
                }
            }

            Handle = new Thread(startProcessing);
            Handle.Start();
            timer1.Start();
        }

        /// <summary>
        /// 文件处理的进程.
        /// </summary>
        private void startProcessing()
        {
            FileStream modeFile_Handover = new FileStream(@"template\实物ID交接单.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook workbook_Handover = new XSSFWorkbook(modeFile_Handover);
            modeFile_Handover.Close();
            XSSFSheet sheet_Handover = (XSSFSheet)workbook_Handover.GetSheetAt(0); //获取工作表 

            foreach (string file in fileNames)
            {
                order = read_OrderFromExcel(file);
                serialNumber++;
                if (skinProgressBar1.InvokeRequired)//不同线程为true，所以这里是true
                {
                    BeginInvoke(new Action<int>(target => { skinProgressBar1.Value = target; }), serialNumber - initialSerialNumber + 1);
                }
                Debug.Write("@" + serialNumber + file + "\r\n");
                sheet_Handover = write_HandoverToExcel(sheet_Handover, order);
                write_BoxsignToExcel(order);
            }

            FileStream file_Handover = new FileStream(folderPath + "/" + String.Format("{0:0000}", initialSerialNumber)+"-"+String.Format("{0:0000}", serialNumber) + "实物ID交接单"+ suffix, FileMode.OpenOrCreate);
            workbook_Handover.Write(file_Handover);
            file_Handover.Close();
            workbook_Handover.Close();
            Handle.Abort();
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
            file_order.order_number = sheet.GetRow(1).GetCell(1).ToString();    //读取订单号（B2）
            file_order.order_name = sheet.GetRow(1).GetCell(4).ToString();      //读取变电站或馈线名称（E2）
            file_order.order_shipping_info = sheet.GetRow(3).GetCell(1).ToString() + sheet.GetRow(3).GetCell(2).ToString() + sheet.GetRow(3).GetCell(3).ToString() + sheet.GetRow(3).GetCell(4).ToString();     //读取发货地址及联系人（B4）

            string[] order_number = new string[5];
            IRow[] row = new IRow[5];
            ICell[] cell = new ICell[5];
            for (int i = 0; i < 5; i++)
            {
                row[i] = sheet.GetRow(4 + i);
                cell[i] = row[i].GetCell(1);

                //判断是否需要计算公式
                string num = cell[i].CellType == CellType.Formula ? cell[i].NumericCellValue.ToString() : cell[i].ToString();

                if (num == "")
                {
                    order_number[i] = "0";
                    string info = file_order.order_number + "标签数量存在空值，请检查！";
                    Logger.AddLogToTXT(info, folderPath + "/log.txt");
                }
                else
                {
                    order_number[i] = num;
                }
            }

            file_order.order_sum_number = order_number[0].ToString();
            file_order.order_A_number = order_number[1].ToString();
            file_order.order_B_number = order_number[2].ToString();
            file_order.order_C_number = order_number[3].ToString();
            file_order.order_D_number = order_number[4].ToString();

            //标签数量校验
            try
            {
                if (Convert.ToInt32(order_number[0]) != (
                    Convert.ToInt32(order_number[1])
                     + Convert.ToInt32(order_number[2])
                     + Convert.ToInt32(order_number[3])
                     + Convert.ToInt32(order_number[4])))
                {
                    string info = file_order.order_number + "标签总数与实际A、B、C、D不符，请检查！";
                    Logger.AddLogToTXT(info,folderPath+"/log.txt");
                }
            }
            catch (Exception e)
            {
                string info = file_order.order_number + "标签数量中发现非数字字符" + e.ToString();
                Logger.AddLogToTXT(info, folderPath + "/log.txt");
            }

            int[] type_num = checkId_FromExcel(sheet, file_order.order_number);
            string[] type = new string[4] { "A", "B", "C", "D" };
            for (int i = 0; i < type_num.Length; i++)
            {
                int num = Convert.ToInt32(order_number[i + 1]);
                if (num != type_num[i])
                {
                    string info = string.Format("{0}:{1}标签数量不符，请检查！{2}-{3}", file_order.order_number, type[i],num,type_num[i]);
                    Logger.AddLogToTXT(info, folderPath + "/log.txt");
                }
            }

            workbook.Close();
            return file_order;
        }

        private int[] checkId_FromExcel(ISheet sheet,string sell_number)
        {
            ICell cell = sheet.GetRow(4).GetCell(1);
            //判断是否需要计算公式
            string num = cell.CellType == CellType.Formula ? cell.NumericCellValue.ToString() : cell.ToString();
            int id_count = Convert.ToInt32(num);    //得到标签总数量

            int error_number = 0;
            int[] order_number = new int[4] { 0, 0, 0, 0 };
            for (int i = 0; i < id_count; i++)
            {
                try
                {
                    string type = sheet.GetRow(10 + i).GetCell(1).ToString().Trim();
                    switch (type)
                    {
                        case "A":
                            order_number[0] += 1;
                            break;
                        case "B":
                            order_number[1] += 1;
                            break;
                        case "C":
                            order_number[2] += 1;
                            break;
                        case "D":
                            order_number[3] += 1;
                            break;
                        default:
                            Debug.WriteLine(string.Format("错误型号：{0}", type));
                            break;
                    }

                    string id = sheet.GetRow(10 + i).GetCell(2).ToString();
                    if (check_Id(id) == false)
                    {
                        string info = sell_number + "实物ID序号" + (i + 1) + "可能有误，请检查！";
                        Logger.AddLogToTXT(info, folderPath + "/log.txt");
                        error_number++;
                    }
                }
                catch (Exception)
                {
                    string info = sell_number + "实物ID序号" + (i + 1) + "引发程序异常，请核对！";
                    Logger.AddLogToTXT(info, folderPath + "/log.txt");
                    error_number++;
                }
            }
            return order_number;
        }

        private bool check_Id(string v)
        {
            if (v == null || v.Length != 24)    //验证这个参数是否为空
                return false;                           //是，就返回False
            //使用ASCII码提取数字
            foreach (char c in v)
            {
                if (Convert.ToInt32(c) < 48 || Convert.ToInt32(c) > 57)
                {
                    return false;
                }
            }
            return true;
        }

        private XSSFSheet write_HandoverToExcel(XSSFSheet sheet, Order_t file_order)
        {
            sheet.GetRow(7 + (serialNumber - initialSerialNumber + 1)).GetCell(0).SetCellValue(file_order.order_number);          //填写订单编号（A*）

            string[] order_num = new string[4];
            order_num[0] = file_order.order_A_number;
            order_num[1] = file_order.order_B_number;
            order_num[2] = file_order.order_C_number;
            order_num[3] = file_order.order_D_number;

            for (int i = 0; i < order_num.Length; i++)
            {
                sheet.GetRow(7 + (serialNumber - initialSerialNumber + 1)).GetCell(3+i).SetCellValue(Convert.ToInt32(order_num[i]));    //填写A、B、C、D型标签数量（D*、E*、F*、G*）

                string num = sheet.GetRow(2 + i).GetCell(5).ToString();
                if (num == "")
                {
                    num = "0";
                }
                int count = Convert.ToInt32(num) + Convert.ToInt32(order_num[i]);
                sheet.GetRow(2 + i).GetCell(5).SetCellValue(count);
            }
            

            return sheet;
        }
        private void write_BoxsignToExcel(Order_t order)
        {
            FileStream modeFile_Boxsign = null;
            modeFile_Boxsign = new FileStream(@"template\实物ID箱签.xlsx", FileMode.Open, FileAccess.Read);

            XSSFWorkbook workbook_Boxsign = new XSSFWorkbook(modeFile_Boxsign);
            modeFile_Boxsign.Close();
            XSSFSheet modeSheet_Boxsign = (XSSFSheet)workbook_Boxsign.GetSheetAt(0); //获取工作表
            
            

            modeSheet_Boxsign.GetRow(0).GetCell(1).SetCellValue(String.Format("{0:0000}", serialNumber));     //填写序号(B1)
            modeSheet_Boxsign.GetRow(7).GetCell(1).SetCellValue(Convert.ToInt32(order.order_A_number));       //填写A型标签数量（B8）
            modeSheet_Boxsign.GetRow(8).GetCell(1).SetCellValue(Convert.ToInt32(order.order_B_number));       //填写B型标签数量（B9）
            modeSheet_Boxsign.GetRow(9).GetCell(1).SetCellValue(Convert.ToInt32(order.order_C_number));       //填写C型标签数量（B10）
            modeSheet_Boxsign.GetRow(10).GetCell(1).SetCellValue(Convert.ToInt32(order.order_D_number));      //填写D型标签数量（B11）

            modeSheet_Boxsign.GetRow(2).GetCell(1).SetCellValue(order.order_number);     //填写订单编号（B3）
            modeSheet_Boxsign.GetRow(0).GetCell(3).SetCellValue(order.order_name);       //填写变电站或馈线名称(D1)

            string[] str = { " ", "；" };
            string[] string_split_word = order.order_shipping_info.Split(str, StringSplitOptions.RemoveEmptyEntries);

            if (string_split_word.Length == 4)
            {
                modeSheet_Boxsign.GetRow(4).GetCell(1).SetCellValue(string_split_word[0] + string_split_word[3]);   //填写收货地址及收货公司（B5）
                modeSheet_Boxsign.GetRow(5).GetCell(1).SetCellValue(string_split_word[1]);   //填写联系人(B6)
                modeSheet_Boxsign.GetRow(5).GetCell(3).SetCellValue(string_split_word[2]);   //填写电话(D6)
            }
            else
            {
                string info = order.order_number + "收货信息格式不符，请检查！";
                Logger.AddLogToTXT(info, folderPath + "/log.txt");
            }
            FileStream file_Boxsign = new FileStream(folderPath + "/实物ID箱签/" + string.Format("{0:0000}", serialNumber) + "-" + order.order_number + "-实物ID箱签"+ suffix, FileMode.Create);

            workbook_Boxsign.Write(file_Boxsign);
            file_Boxsign.Close();
            workbook_Boxsign.Close();
        }

        private void labelFiles_Click(object sender, EventArgs e)
        {
            //选择要处理的文件夹
            FolderBrowserDialog loadFolder = new FolderBrowserDialog();
            loadFolder.Description = "请选择文件夹路径";
            loadFolder.ShowNewFolderButton = false;
            if (loadFolder.ShowDialog() == DialogResult.OK)
            {
                folderPath = loadFolder.SelectedPath;
                labelFiles.Text = folderPath;
                btnStart.Enabled = true;
            }
        }

        public class FileNameSort : IComparer
        {
            //调用DLL
            [System.Runtime.InteropServices.DllImport("Shlwapi.dll", CharSet = CharSet.Unicode)]
            private static extern int StrCmpLogicalW(string param1, string param2);
            
            //前后文件名进行比较。
            public int Compare(object name1, object name2)
            {
                if (null == name1 && null == name2)
                {
                    return 0;
                }
                if (null == name1)
                {
                    return -1;
                }
                if (null == name2)
                {
                    return 1;
                }
                return StrCmpLogicalW(name1.ToString(), name2.ToString());
            }
        }

        int count = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            count ++;
            labelTotalTime.Text = "当前进度：" + (serialNumber - initialSerialNumber +1).ToString() + "/" + fileNames.Length + " 用时：" + count + "s";
            if ((serialNumber - initialSerialNumber + 1) == fileNames.Length)
            {
                timer1.Stop();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            suffix = comboBox1.SelectedItem.ToString();
        }
    }
}
