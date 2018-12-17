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
        private int initialSerialNumber = 1;
        new Thread Handle;   //处理excel的线程
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
            subPath[0] = folderPath + "/实物ID箱签(4个信息)/";
            subPath[1] = folderPath + "/实物ID箱签(3个信息)/";
            foreach (string item in subPath)
            {
                if (false == System.IO.Directory.Exists(item))
                {
                    //创建实物ID箱签文件夹
                    System.IO.Directory.CreateDirectory(item);
                }
            }

            Handle = new Thread(startProcessing);
            Handle.Start();
            timer1.Start();
        }

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

            FileStream file_Handover = new FileStream(folderPath + "/" + String.Format("{0:0000}", serialNumber) + "实物ID交接单.xlsx", FileMode.OpenOrCreate);
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

            string[] num_temp = new string[5];
            IRow[] row = new IRow[5];
            ICell[] cell = new ICell[5];
            for (int i = 0; i < 5; i++)
            {
                row[i] = sheet.GetRow(4 + i);
                cell[i] = row[i].GetCell(1);
                if (cell[i].CellType == CellType.Formula)   //判断是否需要计算公式
                {
                    num_temp[i] = cell[i].NumericCellValue.ToString();
                }
                else
                {
                    num_temp[i] = cell[i].ToString();
                }
            }

            file_order.order_sum_number = num_temp[0].ToString();
            file_order.order_A_number = num_temp[1].ToString();
            file_order.order_B_number = num_temp[2].ToString();
            file_order.order_C_number = num_temp[3].ToString();
            file_order.order_D_number = num_temp[4].ToString();
            
            workbook.Close();
            return file_order;
        }

        private XSSFSheet write_HandoverToExcel(XSSFSheet sheet, Order_t file_order)
        {
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(0).SetCellValue(String.Format("{0:0000}", serialNumber));     //填写序号（A*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(1).SetCellValue(file_order.order_number);          //填写订单编号（B*）

            string[] str = { " ", "；" };
            string[] string_split_word = order.order_shipping_info.Split(str, StringSplitOptions.RemoveEmptyEntries);
            if (string_split_word.Length == 4)
            {
                sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(2).SetCellValue(string_split_word[1]);                   //填写收货人姓名（C*）
            }
            else if (string_split_word.Length == 3) //无收货单位
            {
                sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(2).SetCellValue(string_split_word[1]);                   //填写收货人姓名（C*）
            }
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(3).SetCellValue("A型标签");        //填写A型标签文本（D*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(5).SetCellValue(file_order.order_A_number);        //填写A型标签数量（F*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 1).GetCell(3).SetCellValue("B型标签");               //填写B型标签文本（D*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 1).GetCell(5).SetCellValue(file_order.order_B_number);    //填写B型标签数量（F*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 2).GetCell(3).SetCellValue("C型标签");               //填写C型标签文本（D*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 2).GetCell(5).SetCellValue(file_order.order_C_number);    //填写C型标签数量（F*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 3).GetCell(3).SetCellValue("D型标签");              //填写D型标签文本（D*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1) + 3).GetCell(5).SetCellValue(file_order.order_D_number);              //填写D型标签数量（F*）
            sheet.GetRow(2 + 4 * (serialNumber - initialSerialNumber + 1)).GetCell(4).SetCellValue("1/1");     //填写箱号（E*）
            return sheet;
        }
        private void write_BoxsignToExcel(Order_t order)
        {
            write_BoxsignToExcel_template(0, order);    //0：对应箱贴的模板0
            write_BoxsignToExcel_template(1, order);    //1：对应箱贴的模板1
        }

        private void write_BoxsignToExcel_template(int boxsign_template, Order_t order)
        {
            FileStream modeFile_Boxsign = null;
            if (boxsign_template == 0)
            {
                modeFile_Boxsign = new FileStream(@"template\实物ID箱签0.xlsx", FileMode.Open, FileAccess.Read);
            }
            else if (boxsign_template == 1)
            {
                modeFile_Boxsign = new FileStream(@"template\实物ID箱签1.xlsx", FileMode.Open, FileAccess.Read);
            }
            
            XSSFWorkbook workbook_Boxsign = new XSSFWorkbook(modeFile_Boxsign);
            modeFile_Boxsign.Close();
            XSSFSheet modeSheet_Boxsign = (XSSFSheet)workbook_Boxsign.GetSheetAt(0); //获取工作表
            
            string[] str = { " ", "；" };
            string[] string_split_word = order.order_shipping_info.Split(str, StringSplitOptions.RemoveEmptyEntries);

            modeSheet_Boxsign.GetRow(0).GetCell(1).SetCellValue(String.Format("{0:0000}", serialNumber));     //填写序号(B1)
            modeSheet_Boxsign.GetRow(7).GetCell(1).SetCellValue(order.order_A_number);       //填写A型标签数量（B8）
            modeSheet_Boxsign.GetRow(8).GetCell(1).SetCellValue(order.order_B_number);       //填写B型标签数量（B9）
            modeSheet_Boxsign.GetRow(9).GetCell(1).SetCellValue(order.order_C_number);       //填写C型标签数量（B10）
            modeSheet_Boxsign.GetRow(10).GetCell(1).SetCellValue(order.order_D_number);      //填写D型标签数量（B11）

            if (boxsign_template == 0)
            {
                modeSheet_Boxsign.GetRow(1).GetCell(1).SetCellValue(order.order_number);     //填写订单编号（B2）
                modeSheet_Boxsign.GetRow(0).GetCell(4).SetCellValue(order.order_name);       //填写变电站或馈线名称(E1)

                modeSheet_Boxsign.GetRow(2).GetCell(1).SetCellValue(string_split_word[0]);   //填写收货公司（B3）
                modeSheet_Boxsign.GetRow(4).GetCell(1).SetCellValue(string_split_word[3]);   //填写收货地址（B5）
                modeSheet_Boxsign.GetRow(5).GetCell(1).SetCellValue(string_split_word[1]);   //填写联系人(B6)
                modeSheet_Boxsign.GetRow(5).GetCell(4).SetCellValue(string_split_word[2]);   //填写电话(E6)
            }
            else if (boxsign_template == 1)
            {
                modeSheet_Boxsign.GetRow(2).GetCell(1).SetCellValue(order.order_number);     //填写订单编号（B3）
                modeSheet_Boxsign.GetRow(0).GetCell(3).SetCellValue(order.order_name);       //填写变电站或馈线名称(D1)

                modeSheet_Boxsign.GetRow(4).GetCell(1).SetCellValue(string_split_word[0] + string_split_word[3]);   //填写收货地址及收货公司（B5）
                modeSheet_Boxsign.GetRow(5).GetCell(1).SetCellValue(string_split_word[1]);   //填写联系人(B6)
                modeSheet_Boxsign.GetRow(5).GetCell(3).SetCellValue(string_split_word[2]);   //填写电话(D6)
            }
            
            FileStream file_Boxsign = null;
            if (boxsign_template == 0)
            {
                file_Boxsign = new FileStream(folderPath + "/实物ID箱签(4个信息)/ " + String.Format("{0:0000}", serialNumber) + "-" + order.order_number + "-实物ID箱签.xlsx", FileMode.Create);
            }
            else if (boxsign_template == 1)
            {
                file_Boxsign = new FileStream(folderPath + "/实物ID箱签(3个信息)/ " + String.Format("{0:0000}", serialNumber) + "-" + order.order_number + "-实物ID箱签.xlsx", FileMode.Create);
            }
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
    }
}
