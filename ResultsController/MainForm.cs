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
           
        }
    }
}
