using System;

namespace Excel操作
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.operate2 = new System.Windows.Forms.Button();
            this.myDataSet = new System.Data.DataSet();
            this.OPBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.PictureBox1 = new System.Windows.Forms.PictureBox();
            this.Label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SelectFile = new System.Windows.Forms.Button();
            this.FileName = new System.Windows.Forms.Label();
            this.symbolbutton = new System.Windows.Forms.Button();
            this.offsetBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.片区Box = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.myDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // operate2
            // 
            this.operate2.Location = new System.Drawing.Point(588, 154);
            this.operate2.Name = "operate2";
            this.operate2.Size = new System.Drawing.Size(78, 42);
            this.operate2.TabIndex = 1;
            this.operate2.Text = "Iolist转换按钮";
            this.operate2.UseVisualStyleBackColor = true;
            this.operate2.Click += new System.EventHandler(this.operate2_Click);
            // 
            // myDataSet
            // 
            this.myDataSet.DataSetName = "myDataSet";
            // 
            // OPBox
            // 
            this.OPBox.Location = new System.Drawing.Point(682, 74);
            this.OPBox.Name = "OPBox";
            this.OPBox.Size = new System.Drawing.Size(52, 21);
            this.OPBox.TabIndex = 3;
            this.OPBox.Text = "0";
            this.OPBox.TextChanged += new System.EventHandler(this.OPBox_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 315);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(203, 24);
            this.label1.TabIndex = 4;
            this.label1.Text = "说明：表格命名方式为LCP01 LCP03  \r\n复制一个OP01的输入输出表\r\n";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(574, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 20);
            this.label2.TabIndex = 5;
            this.label2.Text = "LCP数量:";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // PictureBox1
            // 
            this.PictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("PictureBox1.Image")));
            this.PictureBox1.Location = new System.Drawing.Point(534, 308);
            this.PictureBox1.Name = "PictureBox1";
            this.PictureBox1.Size = new System.Drawing.Size(254, 66);
            this.PictureBox1.TabIndex = 12;
            this.PictureBox1.TabStop = false;
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(12, 362);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(53, 12);
            this.Label3.TabIndex = 15;
            this.Label3.Text = "20180918";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(71, 362);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(107, 12);
            this.label4.TabIndex = 14;
            this.label4.Text = "电气部出品V1.2 C#";
            // 
            // SelectFile
            // 
            this.SelectFile.Location = new System.Drawing.Point(12, 241);
            this.SelectFile.Name = "SelectFile";
            this.SelectFile.Size = new System.Drawing.Size(75, 23);
            this.SelectFile.TabIndex = 16;
            this.SelectFile.Text = "选择文件";
            this.SelectFile.UseVisualStyleBackColor = true;
            this.SelectFile.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // FileName
            // 
            this.FileName.AutoSize = true;
            this.FileName.Location = new System.Drawing.Point(12, 282);
            this.FileName.Name = "FileName";
            this.FileName.Size = new System.Drawing.Size(53, 12);
            this.FileName.TabIndex = 17;
            this.FileName.Text = "文件地址";
            // 
            // symbolbutton
            // 
            this.symbolbutton.Location = new System.Drawing.Point(588, 222);
            this.symbolbutton.Name = "symbolbutton";
            this.symbolbutton.Size = new System.Drawing.Size(78, 42);
            this.symbolbutton.TabIndex = 18;
            this.symbolbutton.Text = "变量表生成";
            this.symbolbutton.UseVisualStyleBackColor = true;
            this.symbolbutton.Click += new System.EventHandler(this.symbolbutton_Click);
            // 
            // offsetBox
            // 
            this.offsetBox.Location = new System.Drawing.Point(682, 37);
            this.offsetBox.Name = "offsetBox";
            this.offsetBox.Size = new System.Drawing.Size(52, 21);
            this.offsetBox.TabIndex = 19;
            this.offsetBox.Text = "100";
            this.offsetBox.TextChanged += new System.EventHandler(this.offsetBox_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(544, 38);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(119, 20);
            this.label5.TabIndex = 20;
            this.label5.Text = "设备偏移量:";
            this.label5.Click += new System.EventHandler(this.label5_Click_1);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(682, 222);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(106, 42);
            this.button1.TabIndex = 21;
            this.button1.Text = "手动程序生成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_2);
            // 
            // 片区Box
            // 
            this.片区Box.Location = new System.Drawing.Point(682, 108);
            this.片区Box.Name = "片区Box";
            this.片区Box.Size = new System.Drawing.Size(52, 21);
            this.片区Box.TabIndex = 22;
            this.片区Box.Text = "OP01";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(584, 109);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 20);
            this.label6.TabIndex = 23;
            this.label6.Text = "片区:";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(488, 154);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(78, 42);
            this.button2.TabIndex = 24;
            this.button2.Text = "刷新表";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(488, 222);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(78, 42);
            this.button3.TabIndex = 25;
            this.button3.Text = "选中所需";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 383);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.片区Box);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.offsetBox);
            this.Controls.Add(this.symbolbutton);
            this.Controls.Add(this.FileName);
            this.Controls.Add(this.SelectFile);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.PictureBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OPBox);
            this.Controls.Add(this.operate2);
            this.Name = "Form1";
            this.Text = "IOList 转换Wie输入输出表";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.myDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

 

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button operate2;
        private System.Data.DataSet myDataSet;
        private System.Windows.Forms.TextBox OPBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button SelectFile;
        private System.Windows.Forms.Label FileName;
        private System.Windows.Forms.Button symbolbutton;
        private System.Windows.Forms.TextBox offsetBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox 片区Box;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
    }
}

