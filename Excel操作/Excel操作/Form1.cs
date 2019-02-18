using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
namespace Excel操作
{
    public partial class Form1 : Form
    {
        ExcelEdit excelEdit;
        // private System.ComponentModel.Container  components=null;

        public class PublicValue
        {
               public static string[] LcpName = new string[30];
        }


        public Form1()
        {
            excelEdit = new ExcelEdit();
            InitializeComponent();
          //  GetConnect();//
        }

        private void GetConnect()
        {
            string strConnection = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='{0}';Extended Properties='Excel 8.0'";
            //Properties='Excel 8.0;HDR=YES;IMEX=1'; // 1、HDR表示要把第一行作为数据还是作为列名，作为数据用HDR=no，作为列名用HDR=yes；
            // 2、通过IMEX=1来把混合型作为文本型读取，避免null值。
            string strCon = string.Format(strConnection, "F:\\BaiduYunDownload\\Excel操作\\新建 Microsoft Excel 97-2003 工作表.xls");
            OleDbConnection myConn = new OleDbConnection(strCon);
            strCon = "SELECT * FROM  [Sheet1$]";
            myConn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCon, myConn);
            myCommand.Fill(myDataSet,"table1");
            /*旧的DataSet复制到新的DataSet，oldds.Tables[0].Rows[i][j]进行元素级访问
            myCommand.Fill(newds, "Table1");
            for (int i = 0; i < oldds.Tables[0].Rows.Count; i++)
            {
                //在这里不能使用ImportRow方法将一行导入到news中，因为ImportRow将保留原来DataRow的所有设置(DataRowState状态不变)。
               // 在使用ImportRow后newds内有值，但不能更新到Excel中因为所有导入行的DataRowState != Added
            DataRow nrow = newds.Tables["Table1"].NewRow();
                for (int j = 0; j < newds.Tables[0].Columns.Count; j++)
                {
                    nrow[j] = oldds.Tables[0].Rows[i][j];
                }
                newds.Tables["Table1"].Rows.Add(nrow);
            }
            myCommand.Update(newds, "Table1");
            */

            myConn.Close();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //  excelEdit.Open("C:\\Users\\CDXY\\Desktop\\新建 Microsoft Excel 97-2003 工作表.xls");

            //string filePath = "C:\\Users\\cao\\Desktop\\test.xlsx";


        }

        void chk_CheckedChanged(object sender, EventArgs e)
        {
            string filename = (sender as CheckBox).Text; //获取文件名
                                                         //加载文件代码
        }




        private void button1_Click(object sender, EventArgs e)//YES
        {
            //string fName;
            //openFileDialog1.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            //openFileDialog1.Filter = "Excel2003文件|*.xls|Excel文件|*.xlsx";
            //openFileDialog1.RestoreDirectory = false;
            //openFileDialog1.FilterIndex = 1;
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    fName = openFileDialog1.FileName;//全路径
            //    excelEdit.Open(fName);
            //}
            excelEdit.Open("F:\\BaiduYunDownload\\Excel操作\\新建 Microsoft Excel 工作表.xlsx");
        }

                       private void readData_Click(object sender, EventArgs e)
        {
          
                       excelEdit.ReadRangeArray();
        }

        private void writeData_Click(object sender, EventArgs e)  //YES
        {
            for (int i = 1; i < 10; i++)
            {
                for (int j = 1; j < 10; j++)
                {
    
                    excelEdit.app.Cells[i, j] = (i * 10 + j).ToString();//使用Excel.Cells或者Sheet.Cells均可
                 
                }
            }
        }

        private void getCount_Click(object sender, EventArgs e)  //YES
        {
            //获取已用的范围数据
            int rowsCount = excelEdit.ws.UsedRange.Rows.Count;
            int colsCount = excelEdit.ws.UsedRange.Columns.Count;
            excelEdit.Message();
        }

        private void saveBtn_Click(object sender, EventArgs e)//YES
        {
            excelEdit.SaveAs("F:\\BaiduYunDownload\\Excel操作\\example.xlsx");
            excelEdit.Close();
        }

        private void operate1_Click(object sender, EventArgs e)//YES
        {
            excelEdit.operate1();
        }



        public int lcpNameLength(string[] sheetName)//计算有效长度
        {
            int count =0 ;
            for (int i=0; i<sheetName.Length;i++)
            {
                if (sheetName[i] != null)
                {
                    count = count + 1;
                }
                ; }
           
            return count;

        }




        private void operate2_Click(object sender, EventArgs e)//
        {
            
            if (FileName.Text != null &&( FileName.Text.Contains(".xlsx")|| FileName.Text.Contains(".xls")))
            {
                int OpNum = lcpNameLength(PublicValue.LcpName);
             
                int offset = Convert.ToInt16(offsetBox.Text);
                string 片区 = 片区Box.Text;
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.输入输出表(片区,OpNum, offset, FileName.Text,PublicValue.LcpName);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");
                    
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
                    }

        }
        private void operate3_Click(object sender, EventArgs e)//NO
        {
          //  excelEdit.operate3();
        }

        private void OPBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件夹";
            dialog.Filter = "所有文件(*.*)|*.*";
            string Selectfile = null;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Selectfile = dialog.FileName;
                FileName.Text = Selectfile;
            }

        }

        private void symbolbutton_Click(object sender, EventArgs e) //生成符号表
        {
 
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                         
                int OpNum = lcpNameLength(PublicValue.LcpName);                                                          // MessageBox.Show("请确定是否总共有" + OpNum + "组LCP");
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.符号表(OpNum, FileName.Text, PublicValue.LcpName);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");                   
                    //delete
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }  
        }

        private void label5_Click_1(object sender, EventArgs e)
        {

        }

        private void offsetBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_2(object sender, EventArgs e) //生成电机诊断程序
        {
            
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                int OpNum = lcpNameLength(PublicValue.LcpName);    //Convert.ToInt16(OPBox.Text);
                string 片区 = 片区Box.Text;
                // MessageBox.Show("请确定是否总共有" + OpNum + "组LCP");
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.Manual(片区, FileName.Text);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");
                   
                    //delete
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
             
            string[] workName = new string[30];
            
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                workName = GetExcelSheetNames(FileName.Text);
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }
            if (workName == null)
            {
                MessageBox.Show("获取子表失败");
                return;
            }
            string[] ABC = workName;  
            for (int i = 0; i < ABC.Length; i++)
            {
                CheckBox chk = new CheckBox();
                chk.Name = ABC[i];
                chk.Text = ABC[i];
                chk.Visible = true;
                if (i >= 10 && i < 20)
                {
                    chk.Location = new Point(150, 10 + (i - 10) * 20);
                }
                else if(i >= 20)
                   {
                    chk.Location = new Point(300, 10 + (i - 20) * 20);
                }
                else
                {
                    chk.Location = new Point(10, 10 + i * 20); //你根据实际值计算
                }               
                chk.CheckedChanged += chk_CheckedChanged;
                this.Controls.Add(chk);
            }

            #region 获取Excel工作薄中Sheet页(工作表)名集合
            /// <summary> 
            /// 获取Excel工作薄中Sheet页(工作表)名集合
            /// </summary> 
            /// <param name="excelFile">Excel文件名称及路径,EG:C:\Users\JK\Desktop\导入測试.xls</param> 
            /// <returns>Sheet页名称集合</returns> 
        }

        /// <summary>
        /// 获取Excel子表
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static String[] GetExcelSheetNames(string fileName)
        {
            String[] excelSheets = null;
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;
            try
            {
                string connString = string.Empty;
                string FileType = fileName.Substring(fileName.LastIndexOf("."));
                if (FileType == ".xls")
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                       "Data Source=" + fileName + ";Extended Properties=Excel 8.0;";
                else//.xlsx
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                // 创建连接对象 
                objConn = new OleDbConnection(connString);
                // 打开数据库连接 
                objConn.Open();
                // 得到包括数据架构的数据表 
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    MessageBox.Show("获取数据表为空！");
                    return null;
                }
                excelSheets = new String[dt.Rows.Count];
                int i = 0;
                // 加入工作表名称到字符串数组 
                foreach (DataRow row in dt.Rows)
                {
                    string strSheetTableName = row["TABLE_NAME"].ToString();
                    //过滤无效SheetName
                    if (strSheetTableName.Contains("$") && strSheetTableName.Replace("'", "").EndsWith("$"))
                    {
                        excelSheets[i] = strSheetTableName.Substring(0, strSheetTableName.Length - 1);
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // 清理 
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
            return excelSheets;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

            // string[] lcpName = { };
            PublicValue.LcpName = new string[30];
            foreach (Control ctl in this.Controls)
            {
                int i = 0;
                if (ctl is CheckBox)
                {
                    if (((CheckBox)ctl).Checked)
                    {
                        for (i = 0; i < 30; i++)
                        {
                            if (PublicValue.LcpName[i] == null)
                            {
                                PublicValue.LcpName[i] = ((CheckBox)ctl).Text;
                                break;
                            }


                        }


                    }
                }
            }


            OPBox.Text = lcpNameLength(PublicValue.LcpName).ToString();

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_3(object sender, EventArgs e)  //生成输入映射程序
        {
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                int OpNum = lcpNameLength(PublicValue.LcpName);    //Convert.ToInt16(OPBox.Text);
                string 片区 = 片区Box.Text;
                // MessageBox.Show("请确定是否总共有" + OpNum + "组LCP");
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.InputTranfer(片区, FileName.Text);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");

                    //delete
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }

        }

        private void Output_Click(object sender, EventArgs e) //生成输出映射程序
        {
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                int OpNum = lcpNameLength(PublicValue.LcpName);    //Convert.ToInt16(OPBox.Text);
                string 片区 = 片区Box.Text;
                // MessageBox.Show("请确定是否总共有" + OpNum + "组LCP");
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.OutputTranfer(片区, FileName.Text);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");

                    //delete
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }
        }

        private void Status_Click(object sender, EventArgs e)
        {
            if (FileName.Text != null && (FileName.Text.Contains(".xlsx") || FileName.Text.Contains(".xls")))
            {
                int OpNum = lcpNameLength(PublicValue.LcpName);    //Convert.ToInt16(OPBox.Text);
                string 片区 = 片区Box.Text;
                // MessageBox.Show("请确定是否总共有" + OpNum + "组LCP");
                if (MessageBox.Show("请确定是否总共有" + OpNum + "组LCP", "Confirm Message", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    excelEdit.Status(片区, FileName.Text);
                    excelEdit.Close();
                    MessageBox.Show("全部转换成功");

                    //delete
                }
            }
            else
            {
                MessageBox.Show("没有找到相应文件或者没有找到Excel文件，请重新查找文件");
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        #endregion



        /*  private void btn_SetDataSet_Click(object sender, EventArgs e)//YES
        {
            ////   DataGridView1.DataMember = "[Sheet1$]";
            //  DataGridView1.DataSource = myDataSet.Tables[0];

            DataGridView1.DataMember = "table1";
            DataGridView1.DataSource = myDataSet;
        }
        */




    }
}
