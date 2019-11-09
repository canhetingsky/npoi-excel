namespace npoi_excel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnStart = new System.Windows.Forms.Button();
            this.labelFiles = new System.Windows.Forms.Label();
            this.skinProgressBar1 = new CCWin.SkinControl.SkinProgressBar();
            this.labelTotalTime = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.txbSerialNumber = new CCWin.SkinControl.SkinTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnPrintHandover = new System.Windows.Forms.Button();
            this.btnPrintBoxsign = new System.Windows.Forms.Button();
            this.btnPackingList = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnStart.Location = new System.Drawing.Point(18, 20);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 33);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "开始处理";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // labelFiles
            // 
            this.labelFiles.AutoSize = true;
            this.labelFiles.BackColor = System.Drawing.SystemColors.Control;
            this.labelFiles.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelFiles.ForeColor = System.Drawing.Color.DarkRed;
            this.labelFiles.Location = new System.Drawing.Point(6, 77);
            this.labelFiles.Name = "labelFiles";
            this.labelFiles.Size = new System.Drawing.Size(149, 20);
            this.labelFiles.TabIndex = 1;
            this.labelFiles.Text = "单击选择文件夹";
            this.labelFiles.Click += new System.EventHandler(this.labelFiles_Click);
            // 
            // skinProgressBar1
            // 
            this.skinProgressBar1.Back = null;
            this.skinProgressBar1.BackColor = System.Drawing.Color.Transparent;
            this.skinProgressBar1.BarBack = null;
            this.skinProgressBar1.BarRadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinProgressBar1.ForeColor = System.Drawing.Color.Red;
            this.skinProgressBar1.Location = new System.Drawing.Point(6, 75);
            this.skinProgressBar1.Name = "skinProgressBar1";
            this.skinProgressBar1.RadiusStyle = CCWin.SkinClass.RoundStyle.All;
            this.skinProgressBar1.Size = new System.Drawing.Size(204, 23);
            this.skinProgressBar1.TabIndex = 3;
            // 
            // labelTotalTime
            // 
            this.labelTotalTime.AutoSize = true;
            this.labelTotalTime.Location = new System.Drawing.Point(8, 118);
            this.labelTotalTime.Name = "labelTotalTime";
            this.labelTotalTime.Size = new System.Drawing.Size(35, 12);
            this.labelTotalTime.TabIndex = 4;
            this.labelTotalTime.Text = "*****";
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // txbSerialNumber
            // 
            this.txbSerialNumber.BackColor = System.Drawing.Color.Transparent;
            this.txbSerialNumber.DownBack = null;
            this.txbSerialNumber.Icon = null;
            this.txbSerialNumber.IconIsButton = false;
            this.txbSerialNumber.IconMouseState = CCWin.SkinClass.ControlState.Normal;
            this.txbSerialNumber.IsPasswordChat = '\0';
            this.txbSerialNumber.IsSystemPasswordChar = false;
            this.txbSerialNumber.Lines = new string[] {
        "1"};
            this.txbSerialNumber.Location = new System.Drawing.Point(94, 24);
            this.txbSerialNumber.Margin = new System.Windows.Forms.Padding(0);
            this.txbSerialNumber.MaxLength = 32767;
            this.txbSerialNumber.MinimumSize = new System.Drawing.Size(28, 28);
            this.txbSerialNumber.MouseBack = null;
            this.txbSerialNumber.MouseState = CCWin.SkinClass.ControlState.Normal;
            this.txbSerialNumber.Multiline = false;
            this.txbSerialNumber.Name = "txbSerialNumber";
            this.txbSerialNumber.NormlBack = null;
            this.txbSerialNumber.Padding = new System.Windows.Forms.Padding(5);
            this.txbSerialNumber.ReadOnly = false;
            this.txbSerialNumber.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.txbSerialNumber.Size = new System.Drawing.Size(71, 28);
            // 
            // 
            // 
            this.txbSerialNumber.SkinTxt.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.txbSerialNumber.SkinTxt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txbSerialNumber.SkinTxt.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txbSerialNumber.SkinTxt.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.txbSerialNumber.SkinTxt.Location = new System.Drawing.Point(5, 5);
            this.txbSerialNumber.SkinTxt.Name = "BaseText";
            this.txbSerialNumber.SkinTxt.Size = new System.Drawing.Size(61, 26);
            this.txbSerialNumber.SkinTxt.TabIndex = 0;
            this.txbSerialNumber.SkinTxt.Text = "1";
            this.txbSerialNumber.SkinTxt.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.txbSerialNumber.SkinTxt.WaterText = "";
            this.txbSerialNumber.TabIndex = 5;
            this.txbSerialNumber.Text = "1";
            this.txbSerialNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Left;
            this.txbSerialNumber.WaterColor = System.Drawing.Color.FromArgb(((int)(((byte)(127)))), ((int)(((byte)(127)))), ((int)(((byte)(127)))));
            this.txbSerialNumber.WaterText = "";
            this.txbSerialNumber.WordWrap = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(6, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 19);
            this.label1.TabIndex = 6;
            this.label1.Text = "序号起始";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.labelFiles);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txbSerialNumber);
            this.groupBox1.Location = new System.Drawing.Point(12, 23);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(467, 109);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "步骤一:准备";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox2.Controls.Add(this.btnStart);
            this.groupBox2.Controls.Add(this.skinProgressBar1);
            this.groupBox2.Controls.Add(this.labelTotalTime);
            this.groupBox2.Location = new System.Drawing.Point(12, 158);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(224, 148);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "步骤二:开始";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnPrintHandover);
            this.groupBox3.Controls.Add(this.btnPrintBoxsign);
            this.groupBox3.Controls.Add(this.btnPackingList);
            this.groupBox3.Location = new System.Drawing.Point(257, 158);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(222, 148);
            this.groupBox3.TabIndex = 9;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "步骤三:打印";
            // 
            // btnPrintHandover
            // 
            this.btnPrintHandover.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnPrintHandover.Location = new System.Drawing.Point(57, 109);
            this.btnPrintHandover.Name = "btnPrintHandover";
            this.btnPrintHandover.Size = new System.Drawing.Size(117, 33);
            this.btnPrintHandover.TabIndex = 3;
            this.btnPrintHandover.Text = "打印交接单";
            this.btnPrintHandover.UseVisualStyleBackColor = true;
            // 
            // btnPrintBoxsign
            // 
            this.btnPrintBoxsign.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnPrintBoxsign.Location = new System.Drawing.Point(57, 65);
            this.btnPrintBoxsign.Name = "btnPrintBoxsign";
            this.btnPrintBoxsign.Size = new System.Drawing.Size(117, 33);
            this.btnPrintBoxsign.TabIndex = 2;
            this.btnPrintBoxsign.Text = "打印箱贴";
            this.btnPrintBoxsign.UseVisualStyleBackColor = true;
            // 
            // btnPackingList
            // 
            this.btnPackingList.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnPackingList.Location = new System.Drawing.Point(57, 20);
            this.btnPackingList.Name = "btnPackingList";
            this.btnPackingList.Size = new System.Drawing.Size(117, 33);
            this.btnPackingList.TabIndex = 1;
            this.btnPackingList.Text = "打印装箱单";
            this.btnPackingList.UseVisualStyleBackColor = true;
            // 
            // comboBox1
            // 
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            ".xlsx",
            ".xls"});
            this.comboBox1.Location = new System.Drawing.Point(340, 28);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(82, 20);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(192, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(142, 19);
            this.label2.TabIndex = 8;
            this.label2.Text = "生成文件后缀名";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(584, 349);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label labelFiles;
        private CCWin.SkinControl.SkinProgressBar skinProgressBar1;
        private System.Windows.Forms.Label labelTotalTime;
        private System.Windows.Forms.Timer timer1;
        private CCWin.SkinControl.SkinTextBox txbSerialNumber;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnPrintHandover;
        private System.Windows.Forms.Button btnPrintBoxsign;
        private System.Windows.Forms.Button btnPackingList;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Label label2;
    }
}

