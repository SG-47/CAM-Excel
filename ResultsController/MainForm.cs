using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections;

namespace ResultsController
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private List<System.Data.DataTable> dataTables = new List<System.Data.DataTable>();
        System.Data.DataTable tempDataTable = new System.Data.DataTable();
        private List<string> scoreList = new List<string>();
        private Dictionary<string, List<string>> scoreDic = new Dictionary<string, List<string>>();
        /// <summary>
        /// 导入表按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Title = "请选择要导入的Excel文件";
            openFileDialog.Filter = "Excel文件|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                ImportSelectExcel(openFileDialog.FileName, openFileDialog);
            }
            else
            {
                MessageBox.Show("请选择导入的Excel表格");
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            dataTables.Clear();
            comboBox1.SelectedItem = comboBox1.Items[0];
        }
        /// <summary>
        /// 导入表格
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="openFileDialog"></param>
        private void ImportSelectExcel(string fileName,OpenFileDialog openFileDialog)
        {
            string strConn = string.Empty;
            if(fileName.EndsWith(".xls"))
            {
                strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'", fileName);
                //strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source ={0}; Extended Properties = 'Excel 8.0;HDR=Yes;IMEX=1;'", fileName);
            }
            else if(fileName.EndsWith(".xlsx"))
            {
                strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'", fileName);
            }
            //System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(strConn);
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection();
            conn.ConnectionString = strConn;
            conn.Open();
            string strExcel = "select * from [sheet1$]";
            System.Data.OleDb.OleDbDataAdapter dbDataAdapter = new System.Data.OleDb.OleDbDataAdapter(strExcel, strConn);
            System.Data.DataSet dataSet = new DataSet();
            dbDataAdapter.Fill(dataSet, "table1");
            TabPage tabPage = new TabPage();
            tabPage.Name = "Page" + tabCon_Excel.TabPages.Count;
            tabPage.Text = Path.GetFileName(fileName);
            tabCon_Excel.Controls.Add(tabPage);
            tabCon_Excel.SelectedTab = tabPage;
            DataGridView dataGridView = new DataGridView();
            dataGridView.ReadOnly = true;
            dataGridView.Parent = tabPage;
            dataGridView.Size = tabPage.Size;
            System.Data.DataTable dataTable = dataSet.Tables["table1"];
            dataTable.TableName = tabPage.Text;
            dataGridView.DataSource = dataTable;
            dataTables.Add(dataTable);
            dataGridView.AutoGenerateColumns = false;
        }
        /// <summary>
        /// 合并按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            scoreDic.Clear();
            scoreList.Clear();
            tempDataTable.Clear();
            if(dataTables.Count<=0)
            {
                MessageBox.Show("请先导入要合并的Excel表");
                return;
            }
           else if (dataTables.Count <= 1)
            {
                MessageBox.Show("请导入至少两个表");
                return;
            }
            for (int i = 0; i < dataTables.Count; i++)
            {
                string temp1 = dataTables[i].Rows[0]["F2"].ToString();
                string temp2 = string.Empty;
                if (i+1>=dataTables.Count)
                {
                    temp2 = temp1;
                }
                else
                {
                    temp2 = dataTables[i + 1].Rows[0]["F2"].ToString();
                }
                if (comboBox1.Text == comboBox1.Items[0].ToString())
                {
                    if (temp1 != temp2)
                    {
                        MessageBox.Show(string.Format("合并类型有误，请查看表{0},{1}重新选择合并类型", dataTables[i].TableName, dataTables[i +1].TableName));
                        return;
                    }
                   else
                    {
                        tempDataTable.Merge(dataTables[i], true, MissingSchemaAction.AddWithKey);
                    }
                }
                else if (comboBox1.Text == comboBox1.Items[1].ToString())
                {
                    if (i + 1 >= dataTables.Count)
                    {
                        temp2 = "";
                    }
                    if (temp1 != temp2)
                    {
                        //总成绩合并
                        AddTable(tempDataTable, dataTables[i]);
                    }
                    else
                    {
                        MessageBox.Show(string.Format("合并类型有误，请查看表{0},{1}重新选择合并类型", dataTables[i].TableName, dataTables[i + 1].TableName));
                        return;
                    }
            }
                else
                {
                    MessageBox.Show("合并类型有误，请重新选择");
                }
              
            }

            foreach(TabPage tempPage in tabCon_Excel.TabPages)
            {
                if(tempPage.Text=="合并表")
                {
                    MessageBox.Show("已存在合并表，不能重复合并");
                    return;
                }
            }
            if(comboBox1.Text == comboBox1.Items[1].ToString())
            {

                string id = string.Empty;
                string value = string.Empty;
                for(int i=0;i<tempDataTable.Columns.Count;i++)
                {
                   if(tempDataTable.Rows[0][tempDataTable.Columns[i]].ToString()=="总分")
                    {
                        if (i + 1 >= tempDataTable.Columns.Count)
                        {
                            break;
                        }
                        else
                        {
                            value = tempDataTable.Columns[i + 1].ColumnName.ToString();
                            scoreList.Add(value);
                        }
                    }
                }

                int index = tempDataTable.Columns.Count;
                DataColumn dataColumn = new DataColumn();
                dataColumn.ColumnName = "编号";
                tempDataTable.Columns.Add(dataColumn);
                for(int i=1;i<=scoreList.Count;i++)
                {
                    DataColumn dataColumn1 = new DataColumn();
                    dataColumn1.ColumnName = "比赛项目"+i;
                    tempDataTable.Columns.Add(dataColumn1);
                }
                DataColumn column1 = new DataColumn();
                column1.ColumnName = "比赛总成绩";
                tempDataTable.Columns.Add(column1);
                for(int a=0;a<tempDataTable.Rows.Count;a++)
                {
                    tempDataTable.Rows[a]["编号"] = tempDataTable.Rows[a][tempDataTable.Columns[0]];
                    double f = 0;
                    for(int b=0;b<scoreList.Count;b++)
                    {
                        string tempColumnName = "比赛项目" + (b + 1);
                        string temp = tempDataTable.Rows[a][scoreList[b]].ToString();
                        tempDataTable.Rows[a][tempColumnName] = tempDataTable.Rows[a][scoreList[b]];
                        f += Convert.ToDouble(tempDataTable.Rows[a][tempColumnName]);
                    }
                    tempDataTable.Rows[a]["比赛总成绩"] = f;
                }


            }
            TabPage page = new TabPage();
            page.Name = "Page" + tabCon_Excel.TabPages.Count;
            page.Text = "合并表";
            tabCon_Excel.Controls.Add(page);
            tabCon_Excel.SelectedTab = page;
            DataGridView dataGridView = new DataGridView();
            dataGridView.ReadOnly = true;
            dataGridView.Parent = page;
            dataGridView.Size = page.Size;
            dataGridView.DataSource = tempDataTable;
            dataGridView.AutoGenerateColumns = false;
            
        }


        private void AddTable(System.Data.DataTable tempTable,System.Data.DataTable currentTable)
        {
            if(tempTable.Rows.Count<=0)
            {
                tempTable.Merge(currentTable, true, MissingSchemaAction.AddWithKey);
            }
            else
            {
                int index = tempTable.Columns.Count;
                for(int j=0;j< currentTable.Columns.Count;j++)
                {
                    DataColumn dataColumn = new DataColumn();
                    dataColumn.ColumnName = "F" + (tempTable.Columns.Count + 1);
                    dataColumn.Caption = "F" + (tempTable.Columns.Count + 1);
                    tempTable.Columns.Add(dataColumn);
                }
                for(int a=0;a<tempTable.Rows.Count;a++)
                {
                    for(int b=index;b<tempTable.Columns.Count;b++)
                    {
                        tempTable.Rows[a][tempTable.Columns[b]] = GetCurrentTable(currentTable,a,b-index);
                    }
                }
            }
        }

        private object GetCurrentTable(System.Data.DataTable table,int a,int b)
        {
          
             try
            {
                return table.Rows[a][table.Columns[b]];
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + "请检查Excel表是否匹配");
                return null;
            }
        }

        /// <summary>
        /// 删除表按钮的点击
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            if(tabCon_Excel.TabPages.Count<=0)
            {
                MessageBox.Show("当前未有表");
            }
            else
            {
                tabCon_Excel.TabPages.Remove(tabCon_Excel.SelectedTab);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
           if(tempDataTable.Rows.Count<=0)
            {
                MessageBox.Show("尚未合成表，请先合成再进行导出");
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "保存合成表";
            saveFileDialog.Filter = "Excel文件|*.xls;*.xlsx";
            if(saveFileDialog.ShowDialog()== DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook excelWorkBook = excelApp.Workbooks.Add(System.Type.Missing);//创建工作簿（WorkBook：即Excel文件主体本身）
                Worksheet excelSheet = (Worksheet)excelWorkBook.Worksheets[1]; //创建工作表（即Excel里的子表sheet） 1表示在子表sheet1里进行数据导出
               //excelSheet.Cells.NumberFormat = "@";     //  如果数据中存在数字类型 可以让它变文本格式显示
               //将数据导入到工作表的单元格
                for (int i = 0; i < tempDataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < tempDataTable.Columns.Count; j++)
                    {
                        excelSheet.Cells[i + 1, j + 1] = tempDataTable.Rows[i][j].ToString();   //Excel单元格第一个从索引1开始
                    }
                }

                excelWorkBook.SaveAs(saveFileDialog.FileName, XlFileFormat.xlExcel8);  //将其进行保存到指定的路径

                excelWorkBook.Close();
                excelApp.Quit();  //KillAllExcel(excelApp); 释放可能还没释放的进程
            }
           MessageBox.Show("导出成功");

        }
        /// <summary>
        /// 导出Excel表格的方法1
        /// </summary>
        private void ExportExcel1()
        {
            Microsoft.Office.Interop.Excel.Application appexcel = new Microsoft.Office.Interop.Excel.Application();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
             saveFileDialog.Title = "保存合成表";
        
            System.Reflection.Missing miss = System.Reflection.Missing.Value;

            appexcel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook workbookdata;

            Microsoft.Office.Interop.Excel.Worksheet worksheetdata;

            Microsoft.Office.Interop.Excel.Range rangedata;

            //设置对象不可见

            appexcel.Visible = false;

            System.Globalization.CultureInfo currentci = System.Threading.Thread.CurrentThread.CurrentCulture;

            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-us");

            workbookdata = appexcel.Workbooks.Add(miss);

            worksheetdata = (Microsoft.Office.Interop.Excel.Worksheet)workbookdata.Worksheets.Add(miss, miss, miss, miss);

            //给工作表赋名称

            worksheetdata.Name = "saved";

            for (int i = 0; i < tempDataTable.Columns.Count; i++)

            {

                worksheetdata.Cells[1, i + 1] = tempDataTable.Columns[i].ColumnName.ToString();

            }

            //因为第一行已经写了表头，所以所有数据都应该从a2开始

            rangedata = worksheetdata.get_Range("a2", miss);

            Microsoft.Office.Interop.Excel.Range xlrang = null;

            //irowcount为实际行数，最大行

            int irowcount = tempDataTable.Rows.Count;

            int iparstedrow = 0, icurrsize = 0;

            //ieachsize为每次写行的数值，可以自己设置

            int ieachsize = 1000;

            //icolumnaccount为实际列数，最大列数

            int icolumnaccount = tempDataTable.Columns.Count;

            //在内存中声明一个ieachsize×icolumnaccount的数组，ieachsize是每次最大存储的行数，icolumnaccount就是存储的实际列数

            object[,] objval = new object[ieachsize, icolumnaccount];

            icurrsize = ieachsize;





            while (iparstedrow < irowcount)

            {

                if ((irowcount - iparstedrow) < ieachsize)

                    icurrsize = irowcount - iparstedrow;

                //用for循环给数组赋值

                for (int i = 0; i < icurrsize; i++)

                {

                    for (int j = 0; j < icolumnaccount; j++)

                        objval[i, j] = tempDataTable.Rows[i + iparstedrow][j].ToString();

                    System.Windows.Forms.Application.DoEvents();

                }

                string X = "A" + ((int)(iparstedrow + 2)).ToString();

                string col = "";

                if (icolumnaccount <= 26)

                {

                    col = ((char)('A' + icolumnaccount - 1)).ToString() + ((int)(iparstedrow + icurrsize + 1)).ToString();

                }

                else

                {

                    col = ((char)('A' + (icolumnaccount / 26 - 1))).ToString() + ((char)('A' + (icolumnaccount % 26 - 1))).ToString() + ((int)(iparstedrow + icurrsize + 1)).ToString();

                }

                xlrang = worksheetdata.get_Range(X, col);

                // 调用range的value2属性，把内存中的值赋给excel

                xlrang.Value2 = objval;

                iparstedrow = iparstedrow + icurrsize;

            }

            //保存工作表

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlrang);

            xlrang = null;

            //调用方法关闭excel进程

            appexcel.Visible = true;
        }


 

      

        //public void OutputAsExcelFile(System.Data.DataTable dtTable)
        //{
        //    //Microsoft.Office.Interop.Excel.Application m_xlApp = null;
        //    string filePath = "";
        //    SaveFileDialog s = new SaveFileDialog();
        //    s.Title = "保存Excel文件";
        //    s.Filter = "Excel文件(*.xls)|*.xls";
        //    s.FilterIndex = 1;
        //    if (s.ShowDialog() == DialogResult.OK)
        //        filePath = s.FileName;
        //    else
        //        return;
        //    //导出dataTable到Excel  
        //    //dtTable.Columns.Add("原因");
        //    long rowNum = dtTable.Rows.Count;//行数  
        //    int columnNum = dtTable.Columns.Count;//列数 
        //   // String[] numArr = numStr.Split(',');
        //    Microsoft.Office.Interop.Excel.Application m_xlApp = new Microsoft.Office.Interop.Excel.Application();
        //    m_xlApp.DisplayAlerts = false;//不显示更改提示  
        //    m_xlApp.Visible = false;

        //    Microsoft.Office.Interop.Excel.Workbooks workbooks = m_xlApp.Workbooks;
        //    Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
        //    Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1  

        //    try
        //    {
        //        string[,] datas = new string[rowNum + 1, columnNum];
        //        for (int i = 0; i < columnNum; i++) //写入字段  
        //            datas[0, i] = dtTable.Columns[i].Caption;
        //        //Excel.Range range = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]);  
        //        Microsoft.Office.Interop.Excel.Range range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[1, columnNum]];
        //        range.Interior.ColorIndex = 15;//15代表灰色  
        //        range.Font.Bold = true;
        //        range.Font.Size = 10;

        //        //int r = 0;
        //        //for (r = 0; r < numArr.Length - 1; r++)
        //        //{
        //        //    int numRow = int.Parse(numArr[r]) - 1;
        //        //    for (int i = 0; i < columnNum; i++)
        //        //    {
        //        //        object obj;
        //        //        if (i == columnNum - 1)
        //        //        {
        //        //            obj = "价格格式不符合要求。";
        //        //        }
        //        //        else
        //        //        {
        //        //            obj = dtTable.Rows[numRow][dtTable.Columns[i].ToString()];
        //        //        }
        //        //        datas[r + 1, i] = obj == null ? "" : "'" + obj.ToString().Trim();//在obj.ToString()前加单引号是为了防止自动转化格式
        //        //        Console.WriteLine(datas[r + 1, i]);
        //        //    }
        //        //    System.Windows.Forms.Application.DoEvents();
        //        //    //添加进度条  
        //        //}
        //        //Excel.Range fchR = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
        //        Microsoft.Office.Interop.Excel.Range fchR = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];
        //        fchR.Value2 = datas;

        //        worksheet.Columns.EntireColumn.AutoFit();//列宽自适应。  
        //                                                 //worksheet.Name = "dd";  

        //        //m_xlApp.WindowState = Excel.XlWindowState.xlMaximized;  
        //        m_xlApp.Visible = false;

        //        // = worksheet.get_Range(worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]);  
        //        range = m_xlApp.Range[worksheet.Cells[1, 1], worksheet.Cells[rowNum + 1, columnNum]];

        //        //range.Interior.ColorIndex = 15;//15代表灰色  
        //        range.Font.Size = 9;
        //        range.RowHeight = 14.25;
        //        range.Borders.LineStyle = 1;
        //        range.HorizontalAlignment = 1;
        //        workbook.Saved = true;
        //        workbook.SaveCopyAs(filePath);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("导出异常：" + ex.Message, "导出异常", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //    finally
        //    {
        //        EndReport(m_xlApp);
        //    }

        //   // m_xlApp.Workbooks.Close();
        //    //m_xlApp.Workbooks.Application.Quit();
        //    //m_xlApp.Application.Quit();
        //    //m_xlApp.Quit();
        //    MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //}
        private void EndReport(Microsoft.Office.Interop.Excel.Application m_xlApp)
        {
            object missing = System.Reflection.Missing.Value;
            try
            { }
            catch { }
            finally
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Workbooks);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp.Application);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m_xlApp);
                    m_xlApp = null;
                }
                catch { }
                try
                {
                    //清理垃圾进程  
                    this.killProcessThread();
                }
                catch { }
                GC.Collect();
            }
        }

        private void killProcessThread()
        {
            ArrayList myProcess = new ArrayList();
            for (int i = 0; i < myProcess.Count; i++)
            {
                try
                {
                    System.Diagnostics.Process.GetProcessById(int.Parse((string)myProcess[i])).Kill();
                }
                catch { }
            }
        }
    }

}

