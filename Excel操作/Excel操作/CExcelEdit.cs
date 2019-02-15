using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using System.Web;
using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace Excel操作
    {
    /// <SUMMARY>
    /// Microsoft.Office.Interop.ExcelEdit 的摘要说明
    /// </SUMMARY>
    public class ExcelEdit
    {
        public string mFilename;
        public string txtname;
        public Excel.Application app;
        public Excel.Workbooks wbs;
        public Excel.Workbook wb;
        public Excel.Worksheets wss;
        public Excel.Worksheet ws;
        public Excel.Worksheet ws1;
        public Excel.Worksheet ws2;
        public Excel.Workbook wb1;


        private string filePath;
        private string fileSympolPath;
        private object missing = System.Reflection.Missing.Value;
        public ExcelEdit()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
            this.filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输入输出表.xlsx");
            this.fileSympolPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "符号表.xlsx");
            this.txtname = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OP程序.scl");
        }
        public void Create()//创建一个Excel对象
                            //open 与create使用任意一个
        {
            app = new Excel.Application();
            app.Visible = false;
            wbs = app.Workbooks;
            wb = wbs.Add(true);
            ws = (Excel.Worksheet)wb.ActiveSheet;

        }
        public void Open(string FileName)//打开一个Excel文件
        {
            app = new Excel.Application();
            wbs = app.Workbooks;

            wb = wbs.Open(FileName);
            //wb = wbs.Open(FileName, 0, true, 5,"", "", true, Excel.XlPlatform.xlWindows, "t", false, false, 0, true,Type.Missing,Type.Missing);
            //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
            mFilename = FileName;
            ws = (Excel.Worksheet)wb.ActiveSheet;
            // ws = (Excel.Worksheet)wb.Worksheets["Sheet1"];
        }

        public void Message()
        {
            wb.Author = "Perter";
            string myBookAPP = wb.Application.ToString();
            string myBookName = wb.Name.ToString();
            string myBookFullName = wb.FullName.ToString();
            string myBookFileFormat = wb.FileFormat.ToString();
            int sheetsCount = wb.Sheets.Count;

        }
        public void RangOperate()
        {
            Excel.Range r = app.ActiveCell;
            Excel.Range r1 = ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 4]);
            r1 = ws.get_Range("A1:A3", missing);
            r1 = ws.Cells.get_Range("B2", "B" + 5);
            Excel.Range r2 = (Excel.Range)ws.Cells[1, 1];
            Excel.Range r3 = (Excel.Range)ws.Rows[1, missing];
            Excel.Range r4 = (Excel.Range)ws.Columns[missing, 5];
            r.Font.Bold = true;
            r.Font.Color = System.Drawing.Color.Yellow.ToArgb();
            r.Cells.Interior.Color = System.Drawing.Color.Red.ToArgb();//背景颜色
            r.Borders.Color = 55;
            r.Borders.Weight = Excel.XlBorderWeight.xlThick;
            r.AddComment("这是第一个单元格");
            r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            r.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
            r1.Merge(true);
            ws.Cells[4, 1] = "=sum(A2:A3)";
            r3.NumberFormat = Excel.XlColumnDataType.xlDMYFormat;
        }



        public Excel.Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            ws = (Excel.Worksheet)wb.Worksheets[SheetName];
            return ws;
        }


        public int sheetLength(string[] sheetName)//计算数据表的有效长度
        {
            int count = 0;
            for (int i = 0; i < sheetName.Length; i++)
            {
                if (sheetName[i] != null)
                {
                    count = count + 1;
                }
                ;
            }

            return count;

        }

        public Excel.Worksheet GetShee1t(Int32 SheetNum)
        //获取一个工作表
        {
            ws = (Excel.Worksheet)wb.Worksheets[SheetNum];
            return ws;
        }
        public Excel.Worksheet AddSheet(string SheetName)
        //添加一个工作表,在第一个位置添加
        {
            Excel.Worksheet s = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);//首位置添加
                                                                                                                           //Excel.Worksheet s = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing,wb.Sheets[1], Type.Missing, Type.Missing);//第一个表后
            s.Name = SheetName;
            return s;
        }

        //  public void DelSheet(string SheetName)//删除一个工作表
        public void DelSheet(String SheetName)//删除一个工作表
        {
            //((Excel.Worksheet)wb.Worksheets[SheetName]).Delete();
            ((Excel.Worksheet)wb.Sheets[SheetName]).Delete();//删除第三个工作表
            //((Excel.Worksheet)wb.Sheets[1]).Delete();//删除第三个工作表
        }
        public Excel.Worksheet ReNameSheet(string OldSheetName, string NewSheetName)//重命名一个工作表一
        {
            Excel.Worksheet s = (Excel.Worksheet)wb.Worksheets[OldSheetName];
            s.Name = NewSheetName;
            return s;
        }

        public Excel.Worksheet ReNameSheet(Excel.Worksheet Sheet, string NewSheetName)//重命名一个工作表二
        {

            Sheet.Name = NewSheetName;

            return Sheet;
        }

        public void SetCellValue(int x, int y, object value)
        //ws：要设值的工作表     X行Y列     value   值
        {
            ws.Cells[x, y] = value;
        }

        //public object GetCellValue( int x, int y)
        ////ws：工作表的名称 X行Y列 value 值
        //{
        //    return ws.get_Range(ws.Cells[x, y], ws.Cells[x, y]).Value2;

        //}
        public void ReadRangeArray()
        {
            //这里只读取两列数据，一定要注意rowsint是否正确，当null.tostring在循环中可能会报错
            int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                                                         //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//得到列数


            //取得数据范围区域 (不包括标题列) 
            Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);   //item


            Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint); //Customer
            object[,] arryItem = (object[,])rng1.Value2;   //get range's value
            object[,] arryCus = (object[,])rng2.Value2;
            //将新值赋给一个数组
            string[,] arry = new string[rowsint - 1, 2];
            for (int i = 1; i <= rowsint - 1; i++)
            {
                //Item_Code列
                arry[i - 1, 0] = arryItem[i, 1].ToString();
                //Customer_Name列
                arry[i - 1, 1] = arryCus[i, 1].ToString();
            }
        }


        public void SetCellProperty(Excel.Worksheet ws, int Startx, int Starty, int Endx, int Endy, int size = 12, string FontName = "宋体", Excel.Constants color = Excel.Constants.xlAutomatic, Excel.Constants HorizontalAlignment = Excel.Constants.xlRight)
        //设置一个单元格的属性   字体，   大小，颜色   ，对齐方式
        {

            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = FontName;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }

        public void SetCellProperty(string wsn, int Startx, int Starty, int Endx, int Endy, int size = 12, string FontName = "宋体", Excel.Constants color = Excel.Constants.xlAutomatic, Excel.Constants HorizontalAlignment = Excel.Constants.xlRight)
        {
            //name = "宋体";
            //size = 12;
            //color = Excel.Constants.xlAutomatic;
            //HorizontalAlignment = Excel.Constants.xlRight;

            Excel.Worksheet ws = GetSheet(wsn);
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Name = FontName;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Size = size;
            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).Font.Color = color;

            ws.get_Range(ws.Cells[Startx, Starty], ws.Cells[Endx, Endy]).HorizontalAlignment = HorizontalAlignment;
        }


        public void UniteCells(Excel.Worksheet ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            ws.get_Range(ws.Cells[x1, y1], ws.Cells[x2, y2]).Merge(Type.Missing);
        }

        public void UniteCells(string ws, int x1, int y1, int x2, int y2)
        //合并单元格
        {
            GetSheet(ws).get_Range(GetSheet(ws).Cells[x1, y1], GetSheet(ws).Cells[x2, y2]).Merge(Type.Missing);

        }


        public void InsertTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格插入到Excel指定工作表的指定位置 为在使用模板时控制格式时使用一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {
                    GetSheet(ws).Cells[startX + i, j + startY] = dt.Rows[i][j].ToString();

                }

            }

        }
        public void InsertTable(System.Data.DataTable dt, Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格插入到Excel指定工作表的指定位置二
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[startX + i, j + startY] = dt.Rows[i][j];

                }

            }

        }


        public void AddTable(System.Data.DataTable dt, string ws, int startX, int startY)
        //将内存中数据表格添加到Excel指定工作表的指定位置一
        {

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    GetSheet(ws).Cells[i + startX, j + startY] = dt.Rows[i][j];

                }

            }

        }
        public void AddTable(System.Data.DataTable dt, Excel.Worksheet ws, int startX, int startY)
        //将内存中数据表格添加到Excel指定工作表的指定位置二
        {


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dt.Columns.Count - 1; j++)
                {

                    ws.Cells[i + startX, j + startY] = dt.Rows[i][j];

                }
            }

        }
        public void InsertPictures(string Filename, string ws)
        //插入图片操作一
        {
            GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
            //后面的数字表示位置
        }

        //public void InsertPictures(string Filename, string ws, int Height, int Width)
        //插入图片操作二
        //{
        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}
        //public void InsertPictures(string Filename, string ws, int left, int top, int Height, int Width)
        //插入图片操作三
        //{

        //    GetSheet(ws).Shapes.AddPicture(Filename, MsoTriState.msoFalse, MsoTriState.msoTrue, 10, 10, 150, 150);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementLeft(left);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).IncrementTop(top);
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Height = Height;
        //    GetSheet(ws).Shapes.get_Range(Type.Missing).Width = Width;
        //}

        public void InsertActiveChart(Excel.XlChartType ChartType, string ws, int DataSourcesX1, int DataSourcesY1, int DataSourcesX2, int DataSourcesY2, Excel.XlRowCol ChartDataType)
        //插入图表操作
        {
            ChartDataType = Excel.XlRowCol.xlColumns;
            wb.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            {
                wb.ActiveChart.ChartType = ChartType;
                wb.ActiveChart.SetSourceData(GetSheet(ws).get_Range(GetSheet(ws).Cells[DataSourcesX1, DataSourcesY1], GetSheet(ws).Cells[DataSourcesX2, DataSourcesY2]), ChartDataType);
                wb.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, ws);
            }
        }
        public bool Save(object FileName)
        //保存文档
        {

            if (mFilename == "")
            {
                return false;
            }
            else
            {
                try
                {
                    wb.Save();
                    return true;
                }

                catch (Exception ex)
                {
                    return false;
                }
            }
        }
        public bool SaveAs(object FileName)
        //文档另存为
        {
            try
            {
                wb.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return true;

            }

            catch (Exception ex)
            {
                return false;

            }
        }
        public void Close()
        //关闭一个Excel对象，销毁对象
        {
            //wb.Save();
            wb.Close(Type.Missing, Type.Missing, Type.Missing);
            wbs.Close();
            app.Quit();
            wb = null;
            wbs = null;
            app = null;
            GC.Collect();
        }

        public void operate1()//创建 写，并另存为
        {
            Create();
            ws.SaveAs("F:\\BaiduYunDownload\\Excel操作\\example.xlsx", missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Close();
        }
        public void 输入输出表(string 片区, int value, int offset, string FileName, string[] sheetName)//打开 写，并另存为
        {
            //Open("F:\\BaiduYunDownload\\Excel操作\\新建 Microsoft Excel 工作表.xlsx");
            Open(FileName);
            //wb = wbs[1];//获取工作簿  
            // wb = wbs.Open("2");
            // ws = (Excel.Worksheet)wb.ActiveSheet;
            // string[]  WorkSheetName = { "LCP01", "LCP02" ,"LCP03", "LCP04", "LCP05", "LCP06", "LCP07", "LCP08", "LCP09", "LCP10", "LCP11"
            // , "LCP12", "LCP13", "LCP14", "LCP15", "LCP16", "LCP17", "LCP18", "LCP19", "LCP20"};


            // string[] WorkSheetName = { "LCP01", "LCP02" ,"LCP03", "LCP04"};
            string[] WorkSheetName = sheetName;
            // string[] WorkSheetName = { "LCP07", "LCP10" ,"LCP11"};
            // string[] WorkSheetName = { "LCP14", "LCP15" };

            int OpNum = value;
            int OpNumTemp = 0;

            int InoutQsb = 6;
            int InOutQF = 7;   //QF信号
            int InoutSbfw = 8;  //正转按钮
            int InoutSbbw = 9;   //反转按钮
            int InoutSbStop = 10;   //停止按钮
            int InoutRun = 11;    //运行信号
            int InoutBRun = 12;   //反转运行信号
            int InoutFault = 13;   //故障输入信号
            int InoutBQ1 = 14;    //光电管1
            int InoutBQ2 = 15;    //光电管2
            int InoutBQ3 = 16;   //光电管3
            int InoutBQ4 = 17;    //光电管4
            int InoutBQ5 = 18;    //光电管5
            int InoutSQ1 = 19;    //接近开关1
            int InoutSQ2 = 20;    //接近开关2
            int InoutSQ3 = 21;    //接近开关3
            int InoutSQ4 = 22;    //接近开关4
            int InoutSA1 = 23;    //安全开关1
            int InoutSA2 = 24;    //安全开关2
            int InoutSA3 = 25;    //安全开关3
            int InoutSA4 = 26;    //安全开关4
            int InoutBQ6 = 27;    //光电管6
            int InoutBQ7 = 28;    //光电管7
            int InoutBQ8 = 29;    //光电管8
            int InoutBQ9 = 30;    //光电管9
            int InoutSQ5 = 31;    //接近开关5
            int InoutSQ6 = 32;    //接近开关6
            int InoutSQ7 = 33;    //接近开关7
            int InoutSQ8 = 34;    //接近开关8
                                  //////*输出信号*//////
            int InoutFw = 37;        //输出正转
            int InoutBw = 38;       //输出反转
            int InoutBrake = 39;    //输出抱闸
            int InoutHL1 = 40;      //输出灯1 
            int InoutHL2 = 41;      //输出灯2
            int InoutHL3 = 42;      //输出灯3
            int InoutYV1 = 43;      //输出电磁阀1
            int InoutYV2 = 44;      //输出电磁阀2
            int InoutYV3 = 45;      //输出电磁阀3
            int InoutYV4 = 46;      //输出电磁阀4
            int InoutReset = 47;      //输出电磁阀4


            int InoutNumMax = 0;
            do
            {
                string IoListNum;
                String IoListSymbol;//
                String IoListAdress;
                String IoListQs = null;
                String ioStype = "Input";
                //int IoListColumn = 4; //列
                int IoListRow = 7;//行
                string InoutNumTemp = null;
                int InoutNumMin = 0;

                string InoutNum;
                int InoutRow = 4;
                //  int InoutQs = 5 ;
                try { ws = (Excel.Worksheet)wb.Worksheets[WorkSheetName[OpNumTemp]]; }
                catch {
                    MessageBox.Show("没有找到相应的LCP");
                }
                //此处应该判断是否有OP，然后重新生成
                try
                { ws1 = (Excel.Worksheet)wb.Worksheets[片区]; }
                catch
                {
                    // Microsoft.Office.Interop.Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    // ws1.Name = 片区;
                    AddSheet(片区);
                    ws1 = (Excel.Worksheet)wb.Worksheets[片区];
                }
                ws1.Cells[1, 1] = "平面号";
                ws1.Cells[1, 2] = "设备偏移量";
                ws1.Cells[1, 3] = "母设备";
                ws1.Cells[1, 4] = "母设备的第几台";
                ws1.Cells[1, 5] = "QS";
                ws1.Cells[1, 6] = "QSB";
                ws1.Cells[1, 7] = "QF";
                ws1.Cells[1, 8] = "SB_FW";
                ws1.Cells[1, 9] = "SB_BW";
                ws1.Cells[1, 10] = "SB_Stop";
                ws1.Cells[1, 11] = "F_Running";
                ws1.Cells[1, 12] = "B_Running";
                ws1.Cells[1, 13] = "Fault";
                ws1.Cells[1, 14] = "BQ1";
                ws1.Cells[1, 15] = "BQ2";
                ws1.Cells[1, 16] = "BQ3";
                ws1.Cells[1, 17] = "BQ4";
                ws1.Cells[1, 18] = "BQ5";
                ws1.Cells[1, 19] = "SQ1";
                ws1.Cells[1, 20] = "SQ2";
                ws1.Cells[1, 21] = "SQ3";
                ws1.Cells[1, 22] = "SQ4";
                ws1.Cells[1, 23] = "SA1";
                ws1.Cells[1, 24] = "SA2";
                ws1.Cells[1, 25] = "SA3";
                ws1.Cells[1, 26] = "SA4";
                ws1.Cells[1, 27] = "BQ6";
                ws1.Cells[1, 28] = "BQ7";
                ws1.Cells[1, 29] = "BQ8";
                ws1.Cells[1, 30] = "BQ9";
                ws1.Cells[1, 31] = "SQ5";
                ws1.Cells[1, 32] = "SQ6";
                ws1.Cells[1, 33] = "SQ7";
                ws1.Cells[1, 34] = "SQ8";
                ws1.Cells[1, 35] = "SA1B";
                ws1.Cells[1, 36] = "SA2B";
                ws1.Cells[1, 37] = "FW";
                ws1.Cells[1, 38] = "BW";
                ws1.Cells[1, 39] = "Brake";
                ws1.Cells[1, 40] = "PL1";
                ws1.Cells[1, 41] = "PL2";
                ws1.Cells[1, 42] = "PL3";
                ws1.Cells[1, 43] = "YV1";
                ws1.Cells[1, 44] = "YV2";
                ws1.Cells[1, 45] = "YV3";
                ws1.Cells[1, 46] = "YV4";
                ws1.Cells[1, 47] = "Reset";
                ws1.Cells[1, 48] = "Run";
                ws1.Cells[1, 49] = "backup1";
                ws1.Cells[1, 50] = "backup2";
                ws1.Cells[3, 1] = "start";

                InoutNum = Convert.ToString(ws1.Cells[InoutRow, 1].Value);
                IoListNum = Convert.ToString(ws.Cells[IoListRow, 4].Value);
                IoListSymbol = ws.Cells[IoListRow, 6].Value;
                IoListAdress = ws.Cells[IoListRow, 9].Value;
                //  Console.WriteLine(ws.Cells[2, 6].Value);

                do
                {
                    WriteToInout:
                    try
                    {
                        int ListNum = Convert.ToInt32(ws.Cells[IoListRow, 4].Value);

                        IoListNum = Convert.ToString(ws.Cells[IoListRow, 4].Value);
                        if (Convert.ToString(ws.Cells[IoListRow, 2].Value) == "Input" || Convert.ToString(ws.Cells[IoListRow, 2].Value) == "Output")
                        {
                            ioStype = Convert.ToString(ws.Cells[IoListRow, 2].Value);
                        }

                        if (IoListNum == "" || IoListNum == "0" || IoListNum == null)
                        { IoListRow = IoListRow + 1; }
                        else
                        {
                            IoListSymbol = ws.Cells[IoListRow, 6].Value;
                            IoListAdress = ws.Cells[IoListRow, 9].Value;

                            if (IoListNum == InoutNum)
                            {
                                if (InoutNumMin == 0)
                                {
                                    InoutNumMin = InoutRow;  //最小行
                                }
                                if (ioStype == "input" || ioStype == "Input")
                                {
                                    if (IoListSymbol == "QSB")
                                    {
                                        ws1.Cells[InoutRow, InoutQsb] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SBFW")
                                    {
                                        ws1.Cells[InoutRow, InoutSbfw] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SBBW")
                                    {
                                        ws1.Cells[InoutRow, InoutSbbw] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SBSTOP")
                                    {
                                        ws1.Cells[InoutRow, InoutSbStop] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "Fault" || IoListSymbol == "VF_Fault" || IoListSymbol == "VF")
                                    {
                                        ws1.Cells[InoutRow, InoutFault] = IoListAdress;
                                    }

                                    /*  else if  (IoListSymbol == "QS")
                                       {
                                           ws1.Cells[InoutRow, InoutQs] = IoListQs;
                                       } */
                                    else if (IoListSymbol == "QF")
                                    {
                                        ws1.Cells[InoutRow, InOutQF] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "Run" || IoListSymbol == "RUN")
                                    {
                                        ws1.Cells[InoutRow, InoutRun] = IoListAdress;
                                        ws1.Cells[InoutRow, InoutBRun] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "FRun" || IoListSymbol == "FRUN")
                                    {
                                        ws1.Cells[InoutRow, InoutBRun] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "BRun" || IoListSymbol == "BRUN")
                                    {
                                        ws1.Cells[InoutRow, InoutBRun] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "BQ1")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ1] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ2")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ2] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ3")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ3] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ4")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ4] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ5")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ5] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ6")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ6] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ7")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ7] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ8")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ8] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BQ9")
                                    {
                                        ws1.Cells[InoutRow, InoutBQ9] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ1")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ1] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ2")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ2] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ3")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ3] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ4")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ4] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ5")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ5] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ6")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ6] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ7")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ7] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SQ8")
                                    {
                                        ws1.Cells[InoutRow, InoutSQ8] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SA1")
                                    {
                                        ws1.Cells[InoutRow, InoutSA1] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SA1")
                                    {
                                        ws1.Cells[InoutRow, InoutSA1] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SA2")
                                    {
                                        ws1.Cells[InoutRow, InoutSA2] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SA3")
                                    {
                                        ws1.Cells[InoutRow, InoutSA3] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "SA4")
                                    {
                                        ws1.Cells[InoutRow, InoutSA4] = IoListAdress;
                                    }
                                }
                                //////////输出///////
                                else if (ioStype == "output" || ioStype == "Output")
                                {
                                    if (IoListSymbol == "FKM" || IoListSymbol == "FVF")
                                    {
                                        ws1.Cells[InoutRow, InoutFw] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "BKM" || IoListSymbol == "BVF")
                                    {
                                        ws1.Cells[InoutRow, InoutBw] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "Brake" || IoListSymbol == "BRAKE")
                                    {
                                        ws1.Cells[InoutRow, InoutBrake] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "PL1")
                                    {
                                        ws1.Cells[InoutRow, InoutHL1] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "PL2")
                                    {
                                        ws1.Cells[InoutRow, InoutHL2] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "PL3")
                                    {
                                        ws1.Cells[InoutRow, InoutHL3] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "YV1" || IoListSymbol == "FYV1")
                                    {
                                        ws1.Cells[InoutRow, InoutYV1] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "YV2" || IoListSymbol == "BYV1")
                                    {
                                        ws1.Cells[InoutRow, InoutYV2] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "YV3")
                                    {
                                        ws1.Cells[InoutRow, InoutYV3] = IoListAdress;
                                    }

                                    else if (IoListSymbol == "YV4")
                                    {
                                        ws1.Cells[InoutRow, InoutYV4] = IoListAdress;
                                    }
                                    else if (IoListSymbol == "Reset" || IoListSymbol == "VF_Reset" || IoListSymbol == "RVF" || IoListSymbol == "RESET" || IoListSymbol == "VF_RESET")
                                    {
                                        ws1.Cells[InoutRow, InoutReset] = IoListAdress;
                                    }
                                }

                            }  //判断IOlist里面的平面号和现在的是否一致
                            else
                            {
                                int inoutRowTemp = 3;
                                InoutNumTemp = Convert.ToString(ws1.Cells[inoutRowTemp, 1].Value);
                                string IolistTemp = Convert.ToString(ws.Cells[IoListRow, 4].Value);

                                do
                                {
                                    if (InoutNumTemp != IolistTemp)
                                    {
                                        inoutRowTemp = inoutRowTemp + 1;

                                        InoutNumTemp = Convert.ToString(ws1.Cells[inoutRowTemp, 1].Value);
                                        IolistTemp = Convert.ToString(ws.Cells[IoListRow, 4].Value);
                                        if (InoutNumTemp == "" || InoutNumTemp == null)
                                        {
                                            InoutRow = inoutRowTemp;
                                            ws1.Cells[InoutRow, 1] = ws.Cells[IoListRow, 4];
                                            goto InoutCon;
                                        }
                                    }
                                    else
                                    {
                                        InoutRow = inoutRowTemp;
                                        goto InoutCon;
                                    }
                                }
                                while (inoutRowTemp < 350);
                                InoutRow = InoutRow + 1;

                                ws1.Cells[InoutRow, 1] = ws.Cells[IoListRow, 4];
                                InoutCon: InoutNum = Convert.ToString(ws1.Cells[InoutRow, 1].Value);
                                //IoListRow = IoListRow + 1;
                                goto WriteToInout;
                            }  //判断IOlist里面的平面号和现在的是否一致 如果不一致，先读取Inout表里面是否有该平面号
                            IoListRow = IoListRow + 1;
                        }
                    }
                    catch {

                        if (Convert.ToString(ws.Cells[IoListRow, 4].Value) != "end")
                        { IoListRow = IoListRow + 1; }
                        else
                        { }
                    }
                } while (Convert.ToString(ws.Cells[IoListRow, 4].Value) != "end" && IoListRow < 200);
                InoutNumMax = InoutRow; //最大行

                // 取QS的值//
                IoListRow = 6;
                IoListQs = null;
                IoListNum = null;
                string IolistName = WorkSheetName[OpNumTemp];
                do
                {

                    IoListNum = Convert.ToString(ws.Cells[IoListRow, 4].Value);
                    IoListSymbol = ws.Cells[IoListRow, 6].Value;
                    //IoListNum == WorkSheetName[OpNumTemp]
                    if (IoListNum != null)
                    {
                        if (IoListNum.Contains("LCP"))
                            IoListQs = IoListSymbol;
                        else
                            IoListRow = IoListRow + 1;
                    }
                    else
                        IoListRow = IoListRow + 1;
                }
                while (IoListQs == null && IoListRow < 200);

                // 将QS的值赋值给相应的站台//
                InoutRow = InoutNumMin;
                InoutNumTemp = Convert.ToString(ws1.Cells[InoutRow, 1].Value);
                if (IoListQs != null && IoListQs != "")
                {
                    do
                    {

                        if ((InoutNumTemp != "" || InoutNumTemp != null))
                        {

                            ws1.Cells[InoutRow, 5] = IoListQs;
                            ws1.Cells[InoutRow, 2] = offset;
                            InoutRow = InoutRow + 1;
                            InoutNumTemp = Convert.ToString(ws1.Cells[InoutRow, 1].Value);
                        }

                    }
                    while (InoutRow <= InoutNumMax && InoutRow < 300);

                }

                MessageBox.Show(WorkSheetName[OpNumTemp] + "成功");
                OpNumTemp = OpNumTemp + 1;

            } while (OpNumTemp < OpNum);
            ws1.Cells[InoutNumMax + 3, 1] = "end";


            /* int delCount = 0;
             do {
                 try
                 {
                     ws2 = (Excel.Worksheet)wb.Worksheets[delCount];
                     if (ws2.Name.ToString() != ws1.Name.ToString())
                     {
                         DelSheet(ws2.Name.ToString());
                     }
                 }
                 catch {
                 }

                 delCount++;
             }
             while(delCount < 30);*/


            Save(FileName);
            //SaveAs(filePath);
            //  Close();      
        }


        public void 符号表(int value, string FileName, string[] sheetName)//打开读，失败 生成符号表
        {
            Open(FileName);

            //wb = wbs[1];//获取工作簿
            //ws = (Excel.Worksheet)wb.ActiveSheet;
            //string[] WorkSheetName = { "LCP01", "LCP02" ,"LCP03", "LCP04", "LCP05", "LCP06", "LCP07", "LCP08", "LCP09", "LCP10", "LCP11"
            //   , "LCP12", "LCP13", "LCP14", "LCP15", "LCP16", "LCP17", "LCP18", "LCP19", "LCP20"};
            //  string[] WorkSheetName = { "LCP01", "LCP02" ,"LCP03", "LCP04"};
            string[] WorkSheetName = sheetName;
            //string[] WorkSheetName = { "LCP07", "LCP10" ,"LCP11"};
            //string[] WorkSheetName = { "LCP14", "LCP15" };



            int OpNum = value;
            int OpNumTemp = 0;


            String symbolName = null; //符号名称  列1
            string symbolPath = "IO"; //符号表名称 列2
            string symbolType = "Bool"; // 符号类型  列3
            string symbolAddress = null; // 符号类型  列4
            String symbolComment = null; //符号注释  列5
            String symbolHmiVisible = "True";//符号HMI可显示  列6
            String symbolHmiAccessible = "True"; //符号可访问  列7
            String symbolHmiWriteable = "True";//符号可写  列8

            int symbolRow = 2;

            do
            {
                string IoListNum = null;
                String IoListSymbol = null;//
                String IoListAdress = null;
                string IolistComment = null;

                String ioStype = "Input";
                //int IoListColumn = 4; //列
                int IoListRow = 7;//行
                string symbolRowTemp = null;


                //  int InoutQs = 5 ;
                //
                ws = (Excel.Worksheet)wb.Worksheets[WorkSheetName[OpNumTemp]];
                //ws = (Excel.Worksheet)wb.Worksheets[0];

                try
                { ws1 = (Excel.Worksheet)wb.Worksheets["PLC Tags"]; }
                catch
                {
                    // Microsoft.Office.Interop.Excel.Worksheet ws1 = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    // ws1.Name = 片区;
                    AddSheet("PLC Tags");
                    ws1 = (Excel.Worksheet)wb.Worksheets["PLC Tags"];
                }

                //初始化表格
                ws1.Cells[1, 1] = "Name";
                ws1.Cells[1, 2] = "Path";
                ws1.Cells[1, 3] = "Data Type";
                ws1.Cells[1, 4] = "Logical Address";
                ws1.Cells[1, 5] = "Comment";
                ws1.Cells[1, 6] = "Hmi Visible";
                ws1.Cells[1, 7] = "Hmi Accessible";
                ws1.Cells[1, 8] = "Hmi Writeable";
                ws1.Cells[1, 9] = "Typeobject ID";
                ws1.Cells[1, 10] = "Version ID";

                IoListNum = Convert.ToString(ws.Cells[IoListRow, 4].Value);

                //  Console.WriteLine(ws.Cells[2, 6].Value);
                do {
                    WriteToInout:
                    IoListNum = Convert.ToString(ws.Cells[IoListRow, 4].Value);
                    if (Convert.ToString(ws.Cells[IoListRow, 2].Value) == "Input" || Convert.ToString(ws.Cells[IoListRow, 2].Value) == "Output")
                    {
                        ioStype = Convert.ToString(ws.Cells[IoListRow, 2].Value);
                    }
                    if (IoListNum == "" || IoListNum == "0" || IoListNum == null)
                    { IoListRow = IoListRow + 1; }
                    else
                    {
                        IoListSymbol = Convert.ToString(ws.Cells[IoListRow, 6].Value);
                        symbolAddress = "%" + ws.Cells[IoListRow, 9].Value;
                        symbolComment = ws.Cells[IoListRow, 5].Value + "_" + IoListNum;
                        if (IoListSymbol != null)

                        {
                            IoListSymbol = IoListSymbol.ToUpper();

                            if (ioStype == "input" || ioStype == "Input")
                            {
                                if (IoListSymbol == "QSB")
                                {
                                    symbolName = "QSB" + IoListNum;
                                }

                                else if (IoListSymbol == "READY")
                                {
                                    symbolName = IoListNum + "Ready";
                                }

                                else if (IoListSymbol.Contains("KA"))
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum;
                                }

                                else if (IoListSymbol == "SB")
                                {
                                    symbolName = "SB_" + IoListNum;
                                }

                                else if (IoListSymbol == "ES")
                                {
                                    symbolName = "ES_" + IoListNum;
                                }


                                else if (IoListSymbol == "SBFW")
                                {
                                    symbolName = "Sbfw" + IoListNum;
                                }
                                else if (IoListSymbol == "SBBW")
                                {
                                    symbolName = "Sbbw" + IoListNum;
                                }
                                else if (IoListSymbol == "SBSTOP")
                                {
                                    symbolName = "SBStop" + IoListNum;
                                }
                                else if (IoListSymbol == "FAULT" || IoListSymbol == "VF_FAULT" || IoListSymbol == "VF")
                                {
                                    symbolName = IoListNum + "_Fault";
                                }

                                else if (IoListSymbol.Contains("QS"))
                                {
                                    // symbolName = IoListSymbol;
                                }
                                else if (IoListSymbol == "QF")
                                {
                                    symbolName = "QF" + IoListNum;
                                }
                                else if (IoListSymbol == "Run" || IoListSymbol == "RUN")
                                {
                                    symbolName = IoListNum + "_Running";
                                }
                                else if (IoListSymbol == "FRun" || IoListSymbol == "FRUN")
                                {
                                    symbolName = IoListNum + "Running";
                                }

                                else if (IoListSymbol == "BRun" || IoListSymbol == "BRUN")
                                {
                                    symbolName = IoListNum + "BRunning";
                                }

                                else if (IoListSymbol.Contains("BQ1"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ2"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ3"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ4"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ5"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ6"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ7"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ8"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("BQ9"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ1"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ2"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ3"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ4"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ5"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ6"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ7"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SQ8"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SA1"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SA2"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SA3"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SA4"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SA5"))
                                {
                                    symbolName = IoListSymbol + "A" + IoListNum;
                                }

                                else if (IoListSymbol.Contains("LCP"))
                                {
                                    symbolName = "QS" + "0" + OpNumTemp + 1;
                                }
                                else if (IoListSymbol.Contains("ES"))
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum;
                                }
                                else if (IoListSymbol.Contains("SBL1"))
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum;
                                }

                                else if (IoListSymbol != null)
                                {

                                    symbolName = IoListNum + "_" + IoListSymbol;
                                }

                            }
                            //////////输出///////
                            else if (ioStype == "output" || ioStype == "Output")
                            {
                                if (IoListSymbol == "FKM" || IoListSymbol == "FVF")
                                {
                                    symbolName = IoListNum + "_FW";
                                }
                                else if (IoListSymbol == "BKM" || IoListSymbol == "BVF")
                                {
                                    symbolName = IoListNum + "_BW";
                                }
                                else if (IoListSymbol == "Brake" || IoListSymbol == "BRAKE")
                                {
                                    symbolName = IoListNum + "_Brake";
                                }

                                else if (IoListSymbol == "HVF")
                                {
                                    symbolName = IoListNum + "_Speed";
                                }
                                else if (IoListSymbol == "PL1")
                                {
                                    symbolName = IoListNum + "_PL1";
                                }
                                else if (IoListSymbol == "PL2")
                                {
                                    symbolName = IoListNum + "_PL2";
                                }

                                else if (IoListSymbol == "PL3")
                                {
                                    symbolName = IoListNum + "_PL3";
                                }
                                else if (IoListSymbol == "YV1" || IoListSymbol == "FYV1")
                                {
                                    symbolName = IoListNum + "_YV1";
                                }

                                else if (IoListSymbol == "YV2" || IoListSymbol == "BYV1")
                                {
                                    symbolName = IoListNum + "_YV2";
                                }

                                else if (IoListSymbol == "YV3")
                                {
                                    symbolName = IoListNum + "_YV3";
                                }

                                else if (IoListSymbol == "YV4")
                                {
                                    symbolName = IoListNum + "_YV4";
                                }
                                else if (IoListSymbol == "VF_RESET" || IoListSymbol == "RVF" || IoListSymbol == "RESET")
                                {
                                    symbolName = IoListNum + "_Reset";
                                }

                                else if (IoListSymbol.Contains("ES"))
                                {
                                    symbolName = "ES_" + IoListNum + "_Dis";
                                }
                                else if (IoListSymbol.Contains("SBL1"))
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum + "_Dis";
                                }
                                else if (IoListSymbol == "SB")
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum + "_Dis";
                                }

                                else if (IoListSymbol == "ENABLE")
                                {
                                    symbolName = IoListNum + "_" + IoListSymbol;
                                }
                                else if (IoListSymbol.Contains("KA"))
                                {
                                    symbolName = IoListSymbol + "_" + IoListNum + "_Out";
                                }

                                else if (IoListSymbol != null)
                                {

                                    symbolName = IoListNum + "_" + IoListSymbol;

                                }

                                //   IoListRow = IoListRow + 1;
                                // goto WriteToInout;

                            }
                            ws1.Cells[symbolRow, 1] = symbolName;
                            ws1.Cells[symbolRow, 2] = symbolPath;
                            ws1.Cells[symbolRow, 3] = symbolType;
                            ws1.Cells[symbolRow, 4] = symbolAddress;
                            ws1.Cells[symbolRow, 5] = symbolComment;
                            ws1.Cells[symbolRow, 6] = symbolHmiVisible;
                            ws1.Cells[symbolRow, 7] = symbolHmiAccessible;
                            ws1.Cells[symbolRow, 8] = symbolHmiWriteable;

                            symbolRow = symbolRow + 1;
                            IoListRow = IoListRow + 1;
                            goto WriteToInout;
                        }

                        IoListRow = IoListRow + 1;


                    }

                } while (Convert.ToString(ws.Cells[IoListRow, 4].Value) != "end" && IoListRow < 200);
                // DelSheet("LCP01");
                // GetSheet(WorkSheetName[OpNumTemp]);
                OpNumTemp = OpNumTemp + 1;

            } while (OpNumTemp < OpNum);
            //Save();
            Save(FileName);
            //SaveAs(fileSympolPath);
            // wb.Close(FileName);

            //Close();

        }

        public void Manual(string 片区, string FileName)//打开读，失败 FileName 为打开的Excel
        {
            Open(FileName);
            if (FileName.Contains("电机"))
            { 
                int MNum = 0;
                int M_runStype = 0;
                int MRow = 5;
                int offset = 0;
                int M1or2 = 0;
                int M2type = 0;
                int ActualmNum = 0;
                string saveNmaeText = 片区 + "_Manual.scl";

                txtname = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, saveNmaeText);
                FileStream fs = new FileStream(txtname, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入

                sw.Write("FUNCTION" + "  " + "\"" + 片区 + "_Manual" + "\"" + ": Void");
                sw.Write("\r\n" + @"{ S7_Optimized_Access:=" + @"'TRUE' }");
                sw.Write("\r\n" + "VERSION: 0.1");
                sw.Write("\r\n");
                sw.Write("\r\n" + "BEGIN");



                //清空缓冲区

                //  int InoutQs = 5 ;

                ws = (Excel.Worksheet)wb.Worksheets[片区];
                //  ws1 = (Excel.Worksheet)wb.Worksheets["PLC Tags"];

                //初始化表格
                string AUTO = ws.Cells[4, 2].Value;

                string FAULT_ACK = ws.Cells[4, 3].Value;
                string MOTO_RES = ws.Cells[4, 4].Value;
                string PART_READY = ws.Cells[4, 5].Value;
                string Manual_FW = ws.Cells[4, 6].Value;
                string Manual_BW = ws.Cells[4, 7].Value;
                string TIME_RES = ws.Cells[4, 8].Value;
                string Fault = ws.Cells[4, 9].Value;
                do
                {
                    MNum = Convert.ToInt16(ws.Cells[MRow, 1].Value);
                    M_runStype = Convert.ToInt16(ws.Cells[MRow, 27].Value);
                    M1or2 = Convert.ToInt16(ws.Cells[MRow, 2].Value);
                    M1or2 = Convert.ToInt16(ws.Cells[MRow, 2].Value);
                    M2type = Convert.ToInt16(ws.Cells[MRow, 4].Value);
                    offset = Convert.ToInt16(ws.Cells[MRow, 5].Value);

                    if (M1or2 == 1)
                    {
                        if (M_runStype == 1)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorStandard" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " M_Next_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " FW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw" + ",");
                            //sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            // sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            sw.Write("\r\n" + " FW_Manual_Button:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " BW_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " Transfer_Enable:=" + "\"" + "False" + "\"" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ",");
                            sw.Write("\r\n" + " BW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL3" + ");");
                          
                        }
                        else if (M_runStype == 2)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorOne_way" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " Actual_M_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " UP_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " DN_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw" + ",");
                            //sw.Write("\r\n" + " UP_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            //sw.Write("\r\n" + " DN_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            sw.Write("\r\n" + " UP_Manual_Button:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " DN_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ");");
                            
                        }
                        else if (M_runStype == 3)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorStandard_UPDN" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " M_Next_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Actual_M_ID :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " FW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            //  sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw"+ ",");
                            // sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw"+ ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + Manual_BW + ",");
                            //  sw.Write("\r\n" + " FW_Manual_Button:=" + Manual_FW + ",");
                            // sw.Write("\r\n" + " BW_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " Transfer_Enable:=" + "\"" + "False" + "\"" + ",");
                            sw.Write("\r\n" + " PRX1A:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].SQ1_高位" + ",");
                            sw.Write("\r\n" + " PRX2A:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].SQ2_低位" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ",");                           sw.Write("\r\n" + " BW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL3" + ");");
                            
                        }
                        MRow++;
                    }
                    else if (M1or2 == 2 || M1or2 == 3 || M1or2 == 4 || M1or2 == 5 || M1or2 == 6)
                    {
                        if (M_runStype == 1)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorStandard" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " M_Next_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " FW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw" + ",");
                            //sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            // sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            sw.Write("\r\n" + " FW_Manual_Button:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " BW_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " Transfer_Enable:=" + "\"" + "False" + "\"" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ",");
                            sw.Write("\r\n" + " BW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL3" + ");");
                         
                        }
                        else if (M_runStype == 2)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorOne_way" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " Actual_M_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " UP_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " DN_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw" + ",");
                            //sw.Write("\r\n" + " UP_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            //sw.Write("\r\n" + " DN_Manualrun_Factor:=" + "\"" + "TRUE" + "\"" + ",");
                            sw.Write("\r\n" + " UP_Manual_Button:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " DN_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ");");
                        

                        }
                        else if (M_runStype == 3)
                        {
                            sw.Write("\r\n" + @"///" + MNum + "运行程序");
                            sw.Write("\r\n" + "\"" + "#YF#MotorStandard_UPDN" + "\"" + "(M_ID:=" + MNum + ",");
                            //  sw.Write("\r\n" + "(M_ID:=" + MNum+",");
                            sw.Write("\r\n" + " M_Next_ID :=" + offset + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + offset + ",");
                            sw.Write("\r\n" + " Actual_M_ID :=" + offset + ",");
                            sw.Write("\r\n" + " Part_Ready :=" + PART_READY + ",");
                            sw.Write("\r\n" + " M_Fault:=" + "\"" + "STA" + "\"" + ".M[" + MNum + "].Fault" + ",");
                            sw.Write("\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" + ",");
                            sw.Write("\r\n" + " FW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Fw" + ",");
                            sw.Write("\r\n" + " BW_Autorun_Factor := " + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].A_Factor_Bw" + ",");
                            sw.Write("\r\n" + " M_QS:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].QS" + ",");
                            //  sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Fw"+ ",");
                            // sw.Write("\r\n" + " FW_Manualrun_Factor:=" + "\"" + "Factor" + "\"" + ".Condition[" + MNum + "].M_Factor_Bw"+ ",");
                            sw.Write("\r\n" + " FW_Manualrun_Factor:=" + Manual_FW + ",");
                            sw.Write("\r\n" + " BW_Manualrun_Factor:=" + Manual_BW + ",");
                            //  sw.Write("\r\n" + " FW_Manual_Button:=" + Manual_FW + ",");
                            // sw.Write("\r\n" + " BW_Manual_Button:=" + Manual_BW + ",");
                            sw.Write("\r\n" + " M_Select := " + "\"" + "STA" + "\"" + ".M[" + MNum + "].Selected" + ",");
                            sw.Write("\r\n" + " Transfer_Enable:=" + "\"" + "False" + "\"" + ",");
                            sw.Write("\r\n" + " PRX1A:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].SQ1_高位" + ",");
                            sw.Write("\r\n" + " PRX2A:=" + "\"" + "Input" + "\"" + ".M[" + MNum + "].SQ2_低位" + ",");
                            sw.Write("\r\n" + " FW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL2" + ",");
                            sw.Write("\r\n" + " BW_Dis => " + "\"" + "Output" + "\"" + ".M[" + MNum + "].PL3" + ");");
                        }                       
                         MRow++; 
                    }


                } while (Convert.ToString(ws.Cells[MRow, 1].Value) != "end" && MRow < 200);
                sw.Write("\r\n" + "END_FUNCTION ");


                //Save(); 

                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();

                //SaveAs(fileSympolPath);
                // wb.Close(FileName);

                //Close();
            }

            else
            {
                MessageBox.Show("请选择打开电机数据表");
            }
        }

        public void InputTranfer(string 片区, string FileName)//打开读，失败 FileName 为打开的Excel
        {
            Open(FileName);
            try
            {
                ws1 = (Excel.Worksheet)wb.Worksheets[片区];

                int MRow = 4;
                int offset = 0;
                int M1or2 = 0;
                int M2type = 0;
                int ActualmNum = 0;
                string saveNmaeText = "Input_" + 片区 + ".scl";

                txtname = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, saveNmaeText);
                FileStream fs = new FileStream(txtname, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入

                sw.Write("FUNCTION_BLOCK" + "  " + "\"" + "InputTransfer_" + 片区 + "\"" );
                sw.Write("\r\n" + @"{ S7_Optimized_Access:=" + @"'TRUE' }");
                sw.Write("\r\n" + "VERSION: 0.1");
                sw.Write("\r\n" + "VAR");

                do
                {
                    try
                    {
                        int M_NOTemp = Convert.ToInt32(ws1.Cells[MRow, 1].Value);
                        string M_NO = Convert.ToString(ws1.Cells[MRow, 1].Value);
                        if (M_NO != null)
                        {
                            sw.Write("\r\n" +"M"+ M_NO + ":" + "\"" + "#YF#InputTransfer" + "\"" + ";");
                        }
                    }
                    catch { }
                    MRow = MRow + 1;
                }
                while (Convert.ToString(ws.Cells[MRow, 1].Value) != "end" && MRow < 200);

                sw.Write("\r\n" + "END_VAR");
                sw.Write("\r\n");
                sw.Write("\r\n" + "BEGIN");
                //清空缓冲区


                // ws = (Excel.Worksheet)wb.Worksheets["PLC Tags"];


                MRow = 4;
                do
                {
                    try
                    {
                        int M_NOTemp = Convert.ToInt32(ws1.Cells[MRow, 1].Value);
                        string M_NO = Convert.ToString(ws1.Cells[MRow, 1].Value);
                        string M_OFFSET = Convert.ToString(ws1.Cells[MRow, 2].Value); ;
                        string Original_M = Convert.ToString(ws1.Cells[MRow, 3].Value); ;
                        string Original_Offset = Convert.ToString(ws1.Cells[MRow, 4].Value); ;
                        string QS = Convert.ToString(ws1.Cells[MRow, 5].Value); ;
                        string QSB = Convert.ToString(ws1.Cells[MRow, 6].Value); ;
                        string QF = Convert.ToString(ws1.Cells[MRow, 7].Value); ;
                        string SB_FW = Convert.ToString(ws1.Cells[MRow, 8].Value); ;
                        string SB_BW = Convert.ToString(ws1.Cells[MRow, 9].Value); ;
                        string SB_Stop = Convert.ToString(ws1.Cells[MRow, 10].Value); ;
                        string F_Running = Convert.ToString(ws1.Cells[MRow, 11].Value); ;
                        string B_Running = Convert.ToString(ws1.Cells[MRow, 12].Value); ;
                        string Fault = Convert.ToString(ws1.Cells[MRow, 13].Value); ;
                        string BQ1 = Convert.ToString(ws1.Cells[MRow, 14].Value); ;
                        string BQ2 = Convert.ToString(ws1.Cells[MRow, 15].Value); ;
                        string BQ3 = Convert.ToString(ws1.Cells[MRow, 16].Value); ;
                        string BQ4 = Convert.ToString(ws1.Cells[MRow, 17].Value); ;
                        string BQ5 = Convert.ToString(ws1.Cells[MRow, 18].Value); ;
                        string SQ1 = Convert.ToString(ws1.Cells[MRow, 19].Value); ;
                        string SQ2 = Convert.ToString(ws1.Cells[MRow, 20].Value); ;
                        string SQ3 = Convert.ToString(ws1.Cells[MRow, 21].Value); ;
                        string SQ4 = Convert.ToString(ws1.Cells[MRow, 22].Value); ;
                        string SA1 = Convert.ToString(ws1.Cells[MRow, 23].Value); ;
                        string SA2 = Convert.ToString(ws1.Cells[MRow, 24].Value); ;
                        string SA3 = Convert.ToString(ws1.Cells[MRow, 25].Value); ;
                        string SA4 = Convert.ToString(ws1.Cells[MRow, 26].Value); ;
                        string BQ6 = Convert.ToString(ws1.Cells[MRow, 27].Value); ;
                        string BQ7 = Convert.ToString(ws1.Cells[MRow, 28].Value); ;
                        string BQ8 = Convert.ToString(ws1.Cells[MRow, 29].Value); ;
                        string BQ9 = Convert.ToString(ws1.Cells[MRow, 30].Value); ;
                        string SQ5 = Convert.ToString(ws1.Cells[MRow, 31].Value); ;
                        string SQ6 = Convert.ToString(ws1.Cells[MRow, 32].Value); ;
                        string SQ7 = Convert.ToString(ws1.Cells[MRow, 33].Value); ;
                        string SQ8 = Convert.ToString(ws1.Cells[MRow, 34].Value); ;
                        string SA1B = Convert.ToString(ws1.Cells[MRow, 35].Value); ;
                        string SB1B = Convert.ToString(ws1.Cells[MRow, 36].Value); ;
                        if (M_NO != null)
                        {
                            sw.Write("\r\n" + @"///" + M_NO + "输入信号映射");
                            sw.Write("\r\n" +"#M"+ M_NO + "(");
                            sw.Write("\r\n" + " M_ID:=" + M_NO + ",");
                            // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET);
                            if (QS != null && QS != "")
                            {
                                if (QS.Contains("i") || QS.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " QS:=" + "%" + QS);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " QS:=" + "\"" + QS + "\"");
                                }
                            }
                            if (QSB != null && QSB != "")
                            {
                                if (QSB.Contains("i") || QSB.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " QSB:=" + "%" + QSB);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " QSB:=" + QSB);
                                }
                            }
                            if (QF != null && QF != "")
                            {
                                if (QF.Contains("i") || QF.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " QF:=" + "%" + QF);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " QF:=" + QF);
                                }
                            }
                            if (SB_FW != null && SB_FW != "")
                            {
                                if (SB_FW.Contains("i") || SB_FW.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SB_FW:=" + "%" + SB_FW);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SB_FW:=" + SB_FW);
                                }
                            }
                            if (SB_BW != null && SB_BW != "")
                            {
                                if (SB_BW.Contains("i") || SB_BW.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SB_BW:=" + "%" + SB_BW);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SB_BW:=" + SB_BW);
                                }
                            }
                            if (SB_Stop != null && SB_Stop != "")
                            {
                                if (SB_Stop.Contains("i") || SB_Stop.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SB_STOP:=" + "%" + SB_Stop);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SB_STOP:=" + SB_Stop);
                                }
                            }
                            if (F_Running != null && F_Running != "")
                            {
                                if (F_Running.Contains("i") || F_Running.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " F_Running:=" + "%" + F_Running);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " F_Running:=" + F_Running);
                                }
                            }
                            if (B_Running != null && B_Running != "")
                            {
                                if (B_Running.Contains("i") || B_Running.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " B_Running:=" + "%" + B_Running);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " B_Running:=" + B_Running);
                                }
                            }
                            if (Fault != null && Fault != "")
                            {
                                if (Fault.Contains("i") || Fault.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " Fault:=" + "%" + Fault);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " Fault:=" + Fault);
                                }
                            }
                            if (BQ1 != null && BQ1 != "")
                            {
                                if (BQ1.Contains("i") || BQ1.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ1:=" + "%" + BQ1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ1:=" + BQ1);
                                }
                            }
                            if (BQ2 != null && BQ2 != "")
                            {
                                if (BQ2.Contains("i") || BQ2.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ2:=" + "%" + BQ2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ2:=" + BQ2);
                                }
                            }
                            if (BQ3 != null && BQ3 != "")
                            {
                                if (BQ3.Contains("i") || BQ3.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ3:=" + "%" + BQ3);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ3:=" + BQ3);
                                }
                            }
                            if (BQ4 != null && BQ4 != "")
                            {
                                if (BQ4.Contains("i") || BQ4.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ4:=" + "%" + BQ4);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ4:=" + BQ4);
                                }
                            }
                            if (BQ5 != null && BQ5 != "")
                            {
                                if (BQ5.Contains("i") || BQ5.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ5:=" + "%" + BQ5);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ5:=" + BQ5);
                                }
                            }
                            if (SQ1 != null && SQ1 != "")
                            {
                                if (SQ1.Contains("i") || SQ1.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ1:=" + "%" + SQ1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ1:=" + SQ1);
                                }
                            }
                            if (SQ2 != null && SQ2 != "")
                            {
                                if (SQ2.Contains("i") || SQ2.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ2:=" + "%" + SQ2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ2:=" + SQ2);
                                }
                            }
                            if (SQ3 != null && SQ3 != "")
                            {
                                if (SQ3.Contains("i") || SQ3.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ3:=" + "%" + SQ3);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ3:=" + SQ3);
                                }
                            }
                            if (SQ4 != null && SQ4 != "")
                            {
                                if (SQ4.Contains("i") || SQ4.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ4:=" + "%" + SQ4);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ4:=" + SQ4);
                                }
                            }
                            if (SA1 != null && SA1 != "")
                            {
                                if (SA1.Contains("i") || SA1.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SA1:=" + "%" + SA1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SA1:=" + SA1);
                                }
                            }
                            if (SA2 != null && SA2 != "")
                            {
                                if (SA2.Contains("i") || SA2.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SA2:=" + "%" + SA2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SA2:=" + SA2);
                                }
                            }
                            if (SA3 != null && SA3 != "")
                            {
                                if (SA3.Contains("i") || SA3.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SA3:=" + "%" + SA3);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SA3:=" + SA3);
                                }
                            }
                            if (SA4 != null && SA4 != "")
                            {
                                if (SA4.Contains("i") || SA4.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SA4:=" + "%" + SA4);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SA4:=" + SA4);
                                }
                            }
                            if (BQ6 != null && BQ6 != "")
                            {
                                if (BQ6.Contains("i") || BQ6.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ6:=" + "%" + BQ6);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ6:=" + BQ6);
                                }
                            }
                            if (BQ7 != null && BQ7 != "")
                            {
                                if (BQ7.Contains("i") || BQ7.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ7:=" + "%" + BQ7);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ7:=" + BQ7);
                                }
                            }
                            if (BQ8 != null && BQ8 != "")
                            {
                                if (BQ8.Contains("i") || BQ8.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ8:=" + "%" + BQ8);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ8:=" + BQ8);
                                }
                            }
                            if (BQ9 != null && BQ9 != "")
                            {
                                if (BQ9.Contains("i") || BQ9.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " BQ9:=" + "%" + BQ9);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BQ9:=" + BQ9);
                                }
                            }
                            if (SQ5 != null && SQ5 != "")
                            {
                                if (SQ5.Contains("i") || SQ5.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ5:=" + "%" + SQ5);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ5:=" + SQ5);
                                }
                            }
                            if (SQ6 != null && SQ6 != "")
                            {
                                if (SQ6.Contains("i") || SQ6.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ6:=" + "%" + SQ6);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ6:=" + SQ6);
                                }
                            }
                            if (SQ7 != null && SQ7 != "")
                            {
                                if (SQ7.Contains("i") || SQ7.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ7:=" + "%" + SQ7);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ7:=" + SQ7);
                                }
                            }
                            if (SQ8 != null && SQ8 != "")
                            {
                                if (SQ8.Contains("i") || SQ8.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SQ8:=" + "%" + SQ8);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SQ8:=" + SQ8);
                                }
                            }
                            if (SA1B != null && SA1B != "")
                            {
                                if (SA1B.Contains("i") || SA1B.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SA1B:=" + "%" + SA1B);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SA1B:=" + SA1B);
                                }
                            }
                            if (SB1B != null && SB1B != "")
                            {
                                if (SB1B.Contains("i") || SB1B.Contains("I"))
                                {
                                    sw.Write("," + "\r\n" + " SB1B:=" + "%" + SB1B);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " SB1B:=" + SB1B);
                                }
                            }
                            sw.Write(");");
                        }
                    }
                    catch
                    { }
                    MRow = MRow + 1;
                } while (Convert.ToString(ws.Cells[MRow, 1].Value) != "end" && MRow < 200);
                sw.Write("\r\n");
                sw.Write("\r\n" + "END_FUNCTION_BLOCK ");

                //Save(); 
                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();

                //SaveAs(fileSympolPath);
                // wb.Close(FileName);

                //Close();
            }
            catch
            {
                MessageBox.Show("请查看该Excel表格是否有" + 片区 + "的输入输出表");
            }

        }

        public void OutputTranfer(string 片区, string FileName)//打开读，失败 FileName 为打开的Excel
        {
            Open(FileName);
            try { 

            ws1 = (Excel.Worksheet)wb.Worksheets[片区];
            int MRow = 4;
           // int offset = 0;
           // int M1or2 = 0;
           // int M2type = 0;
            //int ActualmNum = 0;
            string saveNmaeText = "Output_" + 片区 + ".scl";

            txtname = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, saveNmaeText);
            FileStream fs = new FileStream(txtname, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            //开始写入

            sw.Write("FUNCTION_BLOCK " + "  " + "\"" + "OutputTransfer_" + 片区 + "\"" );
            sw.Write("\r\n" + @"{ S7_Optimized_Access:=" + @"'TRUE' }");
            sw.Write("\r\n" + "VERSION: 0.1");
            sw.Write("\r\n" + "VAR");

            do
            {
                    try
                    {
                        int M_NOTemp = Convert.ToInt32(ws1.Cells[MRow, 1].Value);
                        string M_NO = Convert.ToString(ws1.Cells[MRow, 1].Value);
                        if (M_NO != null)
                        {
                            sw.Write("\r\n" +"M"+ M_NO + ":" + "\"" + "#YF#OutputTransfer" + "\"" + ";");
                        }
                    }
                    catch { }
                MRow = MRow + 1;
            }
            while (Convert.ToString(ws.Cells[MRow, 1].Value) != "end" && MRow < 200);

            sw.Write("\r\n" + "END_VAR");
            sw.Write("\r\n");
            sw.Write("\r\n" + "BEGIN");
            //清空缓冲区

            MRow = 4;
            do
            {
                    try
                    {
                        int M_NOTemp = Convert.ToInt32(ws1.Cells[MRow, 1].Value);
                        string M_NO = Convert.ToString(ws1.Cells[MRow, 1].Value);
                        string M_OFFSET = Convert.ToString(ws1.Cells[MRow, 2].Value);
                        string Original_M = Convert.ToString(ws1.Cells[MRow, 3].Value);
                        string Original_Offset = Convert.ToString(ws1.Cells[MRow, 4].Value);
                        string FW = Convert.ToString(ws1.Cells[MRow, 37].Value); ;
                        string BW = Convert.ToString(ws1.Cells[MRow, 38].Value); ;
                        string Brake = Convert.ToString(ws1.Cells[MRow, 39].Value);
                        string HL1 = Convert.ToString(ws1.Cells[MRow, 40].Value); ;
                        string HL2 = Convert.ToString(ws1.Cells[MRow, 41].Value); ;
                        string HL3 = Convert.ToString(ws1.Cells[MRow, 42].Value); ;
                        string YV1 = Convert.ToString(ws1.Cells[MRow, 43].Value); ;
                        string YV2 = Convert.ToString(ws1.Cells[MRow, 44].Value); ;
                        string YV3 = Convert.ToString(ws1.Cells[MRow, 45].Value); ;
                        string YV4 = Convert.ToString(ws1.Cells[MRow, 46].Value); ;
                        string Reset = Convert.ToString(ws1.Cells[MRow, 47].Value);
                        string Run = Convert.ToString(ws1.Cells[MRow, 48].Value); ;
                        string backup1 = Convert.ToString(ws1.Cells[MRow, 49].Value); ;
                        string backup2 = Convert.ToString(ws1.Cells[MRow, 50].Value); ;
                        if ((M_NO != "") && (M_NO != null))
                        {

                            sw.Write("\r\n" + @"///" + M_NO + "输出信号映射");
                            sw.Write("\r\n" +"#M"+ M_NO + "(");
                            sw.Write("\r\n" + " M_ID:=" + M_NO + ",");
                            // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                            sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET);
                            if (FW != null && FW != "")
                            {
                                if (FW.Contains("q") || FW.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " Fw:=" + "%" + FW);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " Fw:=" + "\"" + FW + "\"");
                                }
                            }
                            if (BW != null && BW != "")
                            {
                                if (BW.Contains("q") || BW.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " BW:=" + "%" + BW);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " BW:=" + BW);
                                }
                            }
                            if (Brake != null && Brake != "")
                            {
                                if (Brake.Contains("q") || Brake.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " Brake:=" + "%" + Brake);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " Brake:=" + Brake);
                                }
                            }
                            if (HL1 != null && HL1 != "")
                            {
                                if (HL1.Contains("q") || HL1.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " HL1:=" + "%" + HL1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " HL1:=" + HL1);
                                }
                            }
                            if (HL2 != null && HL2 != "")
                            {
                                if (HL2.Contains("q") || HL2.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " HL2:=" + "%" + HL2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " HL2:=" + HL2);
                                }
                            }
                            if (HL3 != null && HL3 != "")
                            {
                                if (HL3.Contains("q") || HL3.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " HL3:=" + "%" + HL3);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " HL3:=" + HL3);
                                }
                            }
                            if (YV1 != null && YV1 != "")
                            {
                                if (YV1.Contains("q") || YV1.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " YV1:=" + "%" + YV1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " YV1:=" + YV1);
                                }
                            }
                            if (YV2 != null && YV2 != "")
                            {
                                if (YV2.Contains("q") || YV2.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " YV2:=" + "%" + YV2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " YV2:=" + YV2);
                                }
                            }
                            if (YV3 != null && YV3 != "")
                            {
                                if (YV3.Contains("q") || YV3.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " YV3:=" + "%" + YV3);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " YV3:=" + YV3);
                                }
                            }
                            if (YV4 != null && YV4 != "")
                            {
                                if (YV4.Contains("q") || YV4.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " YV4:=" + "%" + YV4);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " YV4:=" + YV4);
                                }
                            }
                            if (Reset != null && Reset != "")
                            {
                                if (Reset.Contains("q") || Reset.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " Reset:=" + "%" + Reset);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " Reset:=" + Reset);
                                }
                            }
                            if (Run != null && Run != "")
                            {
                                if (Run.Contains("q") || Run.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " Run:=" + "%" + Run);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " Run:=" + Run);
                                }
                            }
                            if (backup1 != null && backup1 != "")
                            {
                                if (backup1.Contains("q") || backup1.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " backup1:=" + "%" + backup1);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " backup1:=" + backup1);
                                }
                            }
                            if (backup2 != null && backup2 != "")
                            {
                                if (backup2.Contains("q") || backup2.Contains("Q"))
                                {
                                    sw.Write("," + "\r\n" + " backup2:=" + "%" + backup2);
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " backup2:=" + backup2);
                                }
                            }
                            sw.Write(");");

                        }
                    }
                    catch
                    { }
                MRow = MRow + 1;
            } while (Convert.ToString(ws.Cells[MRow, 1].Value) != "end" && MRow < 200);
            sw.Write("\r\n");
            sw.Write("\r\n" + "END_FUNCTION_BLOCK ");

            //Save(); 
            sw.Flush();
            //关闭流
            sw.Close();
            fs.Close();

            //SaveAs(fileSympolPath);
            // wb.Close(FileName);

            //Close();

        }
            catch
            {
                MessageBox.Show("请查看该Excel表格是否有" + 片区 + "的输入输出表");
            }
        }
        public void Status(string 片区, string FileName)//打开读，失败 FileName 为打开的Excel
        {
            Open(FileName);
            if (FileName.Contains("电机"))
            {

                int MNum = 0; //设备号
                int MRow = 5;
                string M1Type = "0"; //设备类型
                string offset = "0"; //偏移量
                string M1or2 = "0";    //平面号传递
                string M2Type = "0";   //设备2类型
              //  int ActualmNum = 0;  
                string saveNmaeText = 片区 + "_status.scl";
                //清空缓冲区

                //  int InoutQs = 5 ;
                try
                {
                    ws1 = (Excel.Worksheet)wb.Worksheets[片区];
                    //  ws11 = (Excel.Worksheet)wb.Worksheets["PLC Tags"];
                }
                catch
                {
                    MessageBox.Show("电机数据表中没有"+片区+"数据");
                        }
                //初始化表格
                string AUTO = ws1.Cells[4, 2].Value;

                string FAULT_ACK = ws1.Cells[4, 3].Value;
                string MOTO_RES = ws1.Cells[4, 4].Value;
                string PART_READY = ws1.Cells[4, 5].Value;
                string Manual_FW = ws1.Cells[4, 6].Value;
                string Manual_BW = ws1.Cells[4, 7].Value;
                string TIME_RES = ws1.Cells[4, 8].Value;
                string Fault = ws1.Cells[4, 9].Value;

                txtname = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, saveNmaeText);
                FileStream fs = new FileStream(txtname, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                //开始写入

                sw.Write("FUNCTION_BLOCK " + "  " + "\"" + 片区 + "_status" + "\"" );
                sw.Write("\r\n" + @"{ S7_Optimized_Access:=" + @"'TRUE' }");
                sw.Write("\r\n" + "VERSION: 0.1");
                sw.Write("\r\n" + "VAR");

                do
                {
                    try
                    {
                        
                        string M_NO = Convert.ToString(ws1.Cells[MRow, 1].Value);
                        int M_NOTemp = Convert.ToInt16(ws1.Cells[MRow, 1].Value);
                        M1Type = Convert.ToString(ws1.Cells[MRow, 3].Value);
                        M2Type = Convert.ToString(ws1.Cells[MRow, 26].Value);
                        M1or2 = Convert.ToString(ws1.Cells[MRow, 2].Value);
                        if (M_NO != null)
                        {
                            if (M1or2 == "1")
                            {
                                if (M1Type == "1")
                                { sw.Write("\r\n" +"M"+ M_NO + ":" + "\"" + "#YF#StatusSTHF" + "\"" + ";"); }
                                else if (M1Type == "2")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusVF" + "\"" + ";"); }
                                else if (M1Type == "3")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusST" + "\"" + ";"); }
                                else if (M1Type == "4")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusPHIX" + "\"" + ";"); }
                                else if (M1Type == "5")
                                { sw.Write("\r\n" + "M"+ M_NO + ":" + "\"" + "#YF#StatusIO" + "\"" + ";"); }
                                else if (M1Type == "6")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusIO_VF" + "\"" + ";"); }
                            }

                             else if (M1or2 == "2" || M1or2 == "3" || M1or2 == "4" || M1or2 == "5" || M1or2 == "6")
                            {
                                if (M2Type == "1")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusSTHF" + "\"" + ";"); }
                                else if (M2Type == "2")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusVF" + "\"" + ";"); }
                                else if (M2Type == "3")
                                { sw.Write("\r\n" + "M" + "M" + M_NO + ":" + "\"" + "#YF#StatusST" + "\"" + ";"); }
                                else if (M2Type == "4")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusPHIX" + "\"" + ";"); }
                                else if (M2Type == "5")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusIO" + "\"" + ";"); }
                                else if (M2Type == "6")
                                { sw.Write("\r\n" + "M" + M_NO + ":" + "\"" + "#YF#StatusIO_VF" + "\"" + ";"); }
                            }



                            }
                    }
                    catch { }
                    MRow = MRow + 1;
                }
                while (Convert.ToString(ws1.Cells[MRow, 1].Value) != "end" && MRow < 200);

                sw.Write("\r\n" + "END_VAR");
                sw.Write("\r\n");
                sw.Write("\r\n" + "BEGIN");


               

                MRow = 4;
                do
                {
                    try
                    {

                        MNum = Convert.ToInt16(ws1.Cells[MRow, 1].Value);
                        M1or2 = Convert.ToString(ws1.Cells[MRow, 2].Value);
                        M1Type = Convert.ToString(ws1.Cells[MRow, 3].Value);
                        M2Type = Convert.ToString(ws1.Cells[MRow, 26].Value);
                        offset = Convert.ToString(ws1.Cells[MRow, 5].Value);                     
                        String M_BQ_1or2 = Convert.ToString(ws1.Cells[MRow, 6].Value);
                        String M_NEXTBQ_1or2 = Convert.ToString(ws1.Cells[MRow, 7].Value);
                        String M_AUTO_1or2 = Convert.ToString(ws1.Cells[MRow, 8].Value);
                        String M_BQ = Convert.ToString(ws1.Cells[MRow, 9].Value);
                        String M_NEXTBQ = Convert.ToString(ws1.Cells[MRow, 10].Value);
                        int M2Num = Convert.ToInt16(ws1.Cells[MRow, 11].Value);
                        String M_SBorNO = Convert.ToString(ws1.Cells[MRow, 12].Value);
                        int  M_AD_M = Convert.ToInt16(ws1.Cells[MRow, 13].Value);
                        int  M_AD_C = Convert.ToInt16(ws1.Cells[MRow, 14].Value);
                        int  M2_AD_M = Convert.ToInt16(ws1.Cells[MRow, 24].Value);
                        int  M2_AD_C = Convert.ToInt16(ws1.Cells[MRow, 25].Value);
                        String M_T_1or2 = Convert.ToString(ws1.Cells[MRow, 15].Value);
                        String TIMER1 = Convert.ToString(ws1.Cells[MRow, 16].Value);
                        String T_S1 = Convert.ToString(ws1.Cells[MRow, 17].Value);
                        String TIMER2 = Convert.ToString(ws1.Cells[MRow, 18].Value);
                        String T_S2 = Convert.ToString(ws1.Cells[MRow, 19].Value);
                        String KM_Err_Enable = Convert.ToString(ws1.Cells[MRow, 27].Value);
                        String KM_Err_Timer1 = Convert.ToString(ws1.Cells[MRow, 28].Value);
                        String KM_Err_Timer2 = Convert.ToString(ws1.Cells[MRow, 29].Value);
                        String IVALUE1 = Convert.ToString(ws1.Cells[MRow, 30].Value);
                        String IVALUE2 = Convert.ToString(ws1.Cells[MRow, 31].Value);
                        String VFStatus = Convert.ToString(ws1.Cells[MRow, 32].Value);


                        if (M1or2 == "1")
                        {
                            if (M1Type == "1")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + MNum + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M1Type == "2")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "#M" + MNum + "(");
                                sw.Write("\r\n" +"M"+ " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + MNum + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M1Type == "3")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + MNum + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M1Type == "4")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + MNum + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M1Type == "5")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);                            
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"" );
                                sw.Write("," + "\r\n" + " Sensor:= "+ "\"" + "Input" + "\"" + ".M[" + MNum + "]." +M_BQ+"_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1+ "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");          
                            }
                            else if (M1Type == "6")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + MNum + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                             
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + MNum + "]." + M_BQ + "_前到位");                               
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " VF_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + MNum + "].Work_ID);");
                                sw.Write("\r\n");
                            }                          
                        }
                        else if (M1or2 == "2" || M1or2 == "3"||M1or2 == "4"||M1or2 == "5"||M1or2 == "6")
                        {
                            if (M2Type == "1")
                            {
                                sw.Write("\r\n" + @"///" + MNum +"." +M1or2+"故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M2Type == "2")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "." + M1or2 + "故障诊断程序");
                                sw.Write("\r\n" + "M" + MNum + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M2Type == "3")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "." + M1or2 + "故障诊断程序");
                                sw.Write("\r\n" + "M" + M2Num + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M2Type == "4")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "." + M1or2 + "故障诊断程序");
                                sw.Write("\r\n" + "M" + M2Num + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " Address:= " + (M2Num * M_AD_M + M_AD_C));
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");

                            }
                            else if (M2Type == "5")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "." + M1or2 + "故障诊断程序");
                                sw.Write("\r\n" + "M" + M2Num + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);
                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (KM_Err_Enable == "2")
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " KM_Err_Enable := " + "\"" + "True" + "\"");
                                }
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " KM_Err_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");
                            }
                            else if (M2Type == "6")
                            {
                                sw.Write("\r\n" + @"///" + MNum + "." + M1or2 + "故障诊断程序");
                                sw.Write("\r\n" + "M" + M2Num + "(");
                                sw.Write("\r\n" + " M_ID:=" + M2Num + ",");
                                // sw.Write("\r\n" + " M_ID_Offset :=" + M_OFFSET + ",");
                                sw.Write("\r\n" + " M_ID_Offset :=" + offset);

                                sw.Write("," + "\r\n" + " OP_Mode:= " + "\"" + AUTO + "\"");
                                sw.Write("," + "\r\n" + " Sensor:= " + "\"" + "Input" + "\"" + ".M[" + M2Num + "]." + M_BQ + "_前到位");
                                if (M_T_1or2 == "2")
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "False" + "\"");
                                }
                                else
                                {
                                    sw.Write("," + "\r\n" + " T_Enable := " + "\"" + "True" + "\"");
                                }
                                sw.Write("," + "\r\n" + " T_Time := T#" + T_S1 + "S");
                                sw.Write("," + "\r\n" + " T_Reset := " + "\"" + TIME_RES + "\"");
                                sw.Write("," + "\r\n" + " VF_Reset := " + "\"" + MOTO_RES + "\"");
                                sw.Write("," + "\r\n" + " Job_ID:= " + "\"" + "Info" + "\"" + ".M[" + M2Num + "].Work_ID);");
                                sw.Write("\r\n");
                            }
                        }
                    }
                    catch
                    {
                    }

                    MRow++;
                } while (Convert.ToString(ws1.Cells[MRow, 1].Value) != "end" && MRow < 200);
                sw.Write("\r\n" + "END_FUNCTION_BLOCK ");


                //Save(); 

                sw.Flush();
                //关闭流
                sw.Close();
                fs.Close();

                //SaveAs(fileSympolPath);
                // wb.Close(FileName);

                //Close();
            }

            else
            {
                MessageBox.Show("请选择打开电机数据表");
            }
        }

    }
}

