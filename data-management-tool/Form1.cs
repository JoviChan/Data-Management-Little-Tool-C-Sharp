using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Net;
using System.Diagnostics;
using System.Text.RegularExpressions;

//  姓名_电话_车型_服务内容_服务时间
namespace 无忧车秘神器_xp
{
    public partial class Form1 : Form
    {
        public struct point
        {
            public int x;
            public int y;
        }

        private System.Data.OleDb.OleDbConnection sqlCon;
        private System.Data.OleDb.OleDbCommand sqlCmd;
        private System.Data.OleDb.OleDbDataAdapter sqlAdp;
        private System.Windows.Forms.FontDialog fntdlg;
        private string preMobile, startFlag;  //startFlag设置是否需要提醒
        private DateTime preDate;
        private int preNum, saveFlag;
        private Dictionary<point,int> allKeyWords;
        //记录搜索结果的单元框的坐标，分别是两个表格的
        private List<point> keyPosList;
        private List<point> keyPosList2;
        private point aPoint;
        private int pointIndex, pointIndex2;
        private Dictionary<int, int> changeFlag;    //记录每个编号是否被改变,如果value为2表示要删除

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.sqlCon = new System.Data.OleDb.OleDbConnection();
            this.sqlCmd = new System.Data.OleDb.OleDbCommand();
            this.sqlAdp = new System.Data.OleDb.OleDbDataAdapter();
            this.fntdlg = new System.Windows.Forms.FontDialog();
            allKeyWords = new Dictionary<point,int>();
            keyPosList = new List<point>();
            keyPosList2 = new List<point>();
            aPoint = new point();
            textBox1.Text = "点此搜索";
            textBox1.ForeColor = Color.Silver;
            pointIndex = 0;
            pointIndex2 = 0;
            //判断是否设置了开机提醒
            string file2 = Application.StartupPath + "\\startup.log";
            if (File.Exists(file2))
            {
                try
                {
                    FileStream fr = new FileStream(file2, FileMode.Open);
                    StreamReader sr = new StreamReader(fr);
                    startFlag = sr.ReadLine();
                    sr.Close();
                    if (startFlag == "false") checkBox1.Checked = false;
                    else checkBox1.Checked = true;
                }
                catch
                {
                    MessageBox.Show("startup.log文件有错，请将其删除并重新打开", "错误");
                    Application.Exit();
                }

            }
            //没有startup文件，默认为启动开机提醒
            else
            {
                checkBox1.Checked = true;
                startFlag = "true";
            }

            //数据库
            string files = Application.StartupPath + "\\DataBase.mdb";
            if (File.Exists(files))
            {
                saveFlag = 1;
                string aConnectString = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=";
                aConnectString += files;
                sqlCon.ConnectionString = aConnectString;
                sqlCon.Open();
                sqlCmd.Connection = sqlCon;
                sqlAdp = new OleDbDataAdapter("select * from userData order by 编号 asc", sqlCon);
                DataSet ds = new DataSet();
                sqlAdp.Fill(ds, "userData");
                sqlCon.Close();
                Grid1.ColumnCount = 7;
                int ii = 0;
                foreach (var column in ds.Tables[0].Columns)
                {
                    Grid1.Columns[ii].Name = column.ToString();
                    ii++;
                }
                Grid1.Columns[0].FillWeight = 8;
                Grid1.Columns[1].FillWeight = 10;
                Grid1.Columns[2].FillWeight = 15;
                Grid1.Columns[3].FillWeight = 15;
                Grid1.Columns[4].FillWeight = 25;
                Grid1.Columns[5].FillWeight = 12;
                Grid1.Columns[6].FillWeight = 15;

                Grid2.ColumnCount = 7;
                ii = 0;
                foreach (var column in ds.Tables[0].Columns)
                {
                    Grid2.Columns[ii].Name = column.ToString();
                    ii++;
                }
                Grid2.Columns[0].FillWeight = 8;
                Grid2.Columns[1].FillWeight = 10;
                Grid2.Columns[2].FillWeight = 15;
                Grid2.Columns[3].FillWeight = 15;
                Grid2.Columns[4].FillWeight = 25;
                Grid2.Columns[5].FillWeight = 12;
                Grid2.Columns[6].FillWeight = 15;
                ii = 0;
                changeFlag = new Dictionary<int, int>();
                foreach (DataRow row in ds.Tables[0].Rows)
                {
                    Grid1.Rows.Add();
                    Grid1.Rows[ii].Cells[0].Value = row["编号"];
                    changeFlag.Add(int.Parse(Grid1.Rows[ii].Cells[0].Value.ToString()), 0);
                    Grid1.Rows[ii].Cells[1].Value = row["姓名"];
                    Grid1.Rows[ii].Cells[2].Value = row["电话"];
                    Grid1.Rows[ii].Cells[3].Value = row["车型"];
                    Grid1.Rows[ii].Cells[4].Value = row["服务内容"];
                    DateTime tempDate = DateTime.Parse(row["服务时间"].ToString());
                    Grid1.Rows[ii].Cells[5].Value = tempDate.ToShortDateString();
                    Grid1.Rows[ii].Cells[6].Value = row["备注"];
                    if ((tempDate - DateTime.Now).TotalDays <= 1 && (tempDate - DateTime.Now).TotalDays >= -1)
                    {
                        Grid1.Rows[ii].DefaultCellStyle.BackColor = Color.Red;
                    }
                    ii++;
                }
                Grid2Update();
            }
            else
            {
                MessageBox.Show(this, "数据库文件(" + files + ")不存在，请确认");
                Application.Exit();
                this.Dispose();
                this.Close();
            }
            string file = Application.StartupPath + "\\font.log";
            if (File.Exists(file))
            {
                StreamReader sr = new StreamReader(file, System.Text.Encoding.Unicode);
                string fn = sr.ReadLine();
                float sz = float.Parse(sr.ReadLine());
                Grid1.Font = new Font(fn, sz);
                sr.Close();
            }
        }

        //搜索一个表格
        private void searchGrid(DataGridView grid)
        {
            int firstOne = 0;
            pointIndex = 0;
            for (int i = 0; i < grid.RowCount; i++)
            {
                if (grid.Rows[i].Cells[0] == null) continue;
                for (int j = 0; j < grid.ColumnCount; j++)
                {
                    if (grid[j, i].Style.BackColor == Color.Yellow)
                    {
                        aPoint.x = i;
                        aPoint.y = j;
                        if (allKeyWords[aPoint] == 0) grid[j, i].Style.BackColor = Color.White;
                        else grid[j, i].Style.BackColor = Color.Red;
                    }
                }
            }
            allKeyWords.Clear();
            keyPosList.Clear();
            if (textBox1.Text != "")
            {
                for (int i = 0; i < grid.RowCount - 1; i++)
                {
                    for (int j = 0; j < grid.ColumnCount; j++)
                    {
                        if (grid[j, i].Value.ToString().Contains(textBox1.Text))
                        {
                            //allkeywords第三项判别改变之前该单元格背景是否为红
                            if (grid.Rows[i].DefaultCellStyle.BackColor == Color.Red)
                            {
                                aPoint.x = i;
                                aPoint.y = j;
                                allKeyWords.Add(aPoint, 1);
                            }
                            else
                            {
                                aPoint.x = i;
                                aPoint.y = j;
                                allKeyWords.Add(aPoint, 0);
                            }
                            grid[j, i].Style.BackColor = Color.Yellow;
                            keyPosList.Add(aPoint);
                            if (firstOne == 0)
                            {
                                grid.CurrentCell = grid[j, i];
                                firstOne = 1;
                                grid.DefaultCellStyle.SelectionBackColor = Color.Purple;
                            }
                        }
                    }
                }
            }
            else grid.DefaultCellStyle.SelectionBackColor = Color.Blue;
        }

        //更新表格2
        private void Grid2Update()
        {
            Grid2.Rows.Clear();
            for(int i = 0; i < Grid1.Rows.Count - 1; i++)
            {
                DateTime tempDate;
                if (DateTime.TryParse(Grid1.Rows[i].Cells[5].Value.ToString(), out tempDate))
                {
                    if ((tempDate - DateTime.Now).TotalDays <= 1 && (tempDate - DateTime.Now).TotalDays >= -1)
                    {
                        Grid2.Rows.Add();
                        for (int j = 0; j < 7; j++) Grid2.Rows[Grid2.Rows.Count - 1].Cells[j].Value = Grid1.Rows[i].Cells[j].Value;
                    }
                }    
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (fntdlg.ShowDialog() == DialogResult.OK)
            {
                this.Grid1.Font = fntdlg.Font;
                string file = Application.StartupPath + "\\font.log";
                StreamWriter sw = new StreamWriter(file, false, System.Text.Encoding.Unicode);
                sw.WriteLine(fntdlg.Font.Name);
                sw.WriteLine(fntdlg.Font.Size.ToString());
                sw.Close();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.ForeColor = Color.Black;
            searchGrid(Grid1);
            int firstOne = 0;
            pointIndex2 = 0;
            for (int i = 0; i < Grid2.RowCount; i++)
            {
                if (Grid2.Rows[i].Cells[0] == null) continue;
                for (int j = 0; j < Grid2.ColumnCount; j++)
                {
                    if (Grid2[j, i].Style.BackColor == Color.Yellow)
                    {
                        aPoint.x = i;
                        aPoint.y = j;
                        Grid2[j, i].Style.BackColor = Color.White;
                    }
                }
            }
            keyPosList2.Clear();
            if (textBox1.Text != "")
            {
                for (int i = 0; i < Grid2.RowCount; i++)
                {
                    for (int j = 0; j < Grid2.ColumnCount; j++)
                    {
                        if (Grid2[j, i].Value.ToString().Contains(textBox1.Text))
                        {
                            aPoint.x = i;
                            aPoint.y = j;
                            Grid2[j, i].Style.BackColor = Color.Yellow;
                            keyPosList2.Add(aPoint);
                            if (firstOne == 0)
                            {
                                Grid2.CurrentCell = Grid2[j, i];
                                firstOne = 1;
                                Grid2.DefaultCellStyle.SelectionBackColor = Color.Purple;
                            }
                        }
                    }
                }
            }
            else Grid2.DefaultCellStyle.SelectionBackColor = Color.Blue;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "点此搜索") textBox1.Text = "";
        }

        private void Grid1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (Grid1.CurrentRow.Cells[0].Value == null)
            {
                int tempNum = Grid1.CurrentRow.Index + 1;
                while (changeFlag.ContainsKey(tempNum))
                {
                    if (changeFlag[tempNum] == 2) break;
                    tempNum++;
                } 
                Grid1.CurrentRow.Cells[0].Value = tempNum;
                changeFlag.Add(tempNum, 1);
                if(preNum != 0) changeFlag.Remove(preNum);
                preNum = tempNum;
            }

            if (Grid1.CurrentCell.Value.ToString() == "")
            {
                Grid1.CurrentCell.Value = "";
                if (Grid1.CurrentCell.ColumnIndex == 5) Grid1.CurrentCell.Value = preDate.ToShortDateString();
            } 
            else if (Grid1.CurrentCell.ColumnIndex == 0)
            {
                int tempNum;
                if (!int.TryParse(Grid1.CurrentCell.Value.ToString(), out tempNum))
                {
                    if (preNum != 0) Grid1.CurrentCell.Value = preNum;
                    else
                    {
                        tempNum = Grid1.CurrentRow.Index + 1;
                        while (changeFlag.ContainsKey(tempNum))
                        {
                            if (changeFlag[tempNum] == 2) break;
                            tempNum++;
                        } 
                        Grid1.CurrentRow.Cells[0].Value = tempNum;
                        changeFlag.Add(tempNum, 1);
                    }
                    MessageBox.Show("编号必须为数字且不能和之前的编号重复", "提示");
                }
                else
                {
                    if (changeFlag.ContainsKey(tempNum))
                    {
                        if (tempNum != preNum && changeFlag[tempNum] != 2)
                        {
                            if (preNum != 0) Grid1.CurrentCell.Value = preNum;
                            else
                            {
                                tempNum = Grid1.CurrentRow.Index + 1;
                                while (changeFlag.ContainsKey(tempNum)) tempNum++;
                                Grid1.CurrentRow.Cells[0].Value = tempNum;
                                changeFlag.Add(tempNum, 1);
                            }
                            MessageBox.Show("编号必须为数字且不能和之前的编号重复", "提示");
                        }
                        else if (changeFlag[tempNum] == 2)//之前删掉的可以添加
                        {
                            Grid1.CurrentCell.Value = tempNum;
                            changeFlag[tempNum] = 1;
                            if(preNum != 0)changeFlag.Remove(preNum);
                            saveFlag = 0;
                            preNum = tempNum;
                        }
                    }
                    else if (tempNum == 0)
                    {
                        if (preNum != 0) Grid1.CurrentCell.Value = preNum;
                        else
                        {
                            tempNum = Grid1.CurrentRow.Index + 1;
                            while (changeFlag.ContainsKey(tempNum))
                            {
                                if (changeFlag[tempNum] == 2) break;
                                tempNum++;
                            } 
                            Grid1.CurrentRow.Cells[0].Value = tempNum;
                            changeFlag.Add(tempNum, 1);
                        }
                        MessageBox.Show("编号不能为0", "提示");
                    }
                    else
                    {
                        Grid1.CurrentCell.Value = tempNum;
                        changeFlag.Add(tempNum, 1);
                        if (preNum != 0) changeFlag.Remove(preNum);
                        saveFlag = 0;
                        preNum = tempNum;
                    }             
                }
            }
            else if (Grid1.CurrentCell.ColumnIndex == 2)
            {
                if (!Regex.IsMatch(Grid1.CurrentCell.Value.ToString(), @"\A(\d+-\d+|\d+)(/\d+-\d+|/\d+)*\Z"))
                {
                    Grid1.CurrentCell.Value = preMobile;
                    MessageBox.Show("电话号码格式应为\r\n***-******或者*********\r\n*为数字\n若有多个电话，则以 / 分隔", "提示");
                }
                else
                {
                    saveFlag = 0;
                    changeFlag[int.Parse(Grid1.CurrentRow.Cells[0].Value.ToString())] = 1;
                    preMobile = Grid1.CurrentCell.Value.ToString();
                    Grid1.CurrentCell.Value = preMobile;
                } 
            }
            else if (Grid1.CurrentCell.ColumnIndex == 5)
            {
                if (!Regex.IsMatch(Grid1.CurrentCell.Value.ToString(), @"\A\d+/\d+/\d+\Z"))
                {
                    Grid1.CurrentCell.Value = preDate.ToShortDateString();
                    MessageBox.Show("日期格式应为\r\n年（数字）/月（数字）/日（数字）", "提示");
                }
                else
                {
                    try
                    {
                        saveFlag = 0;
                        changeFlag[int.Parse(Grid1.CurrentRow.Cells[0].Value.ToString())] = 1;
                        preDate = DateTime.Parse(Grid1.CurrentCell.Value.ToString());
                        Grid1.CurrentCell.Value = preDate.ToShortDateString();
                    }
                    catch
                    {
                        MessageBox.Show("不存在的日期！", "警告");
                        Grid1.CurrentCell.Value = preDate.ToShortDateString();
                    }
                }
            }
            else
            {
                Grid1.CurrentCell.Value = Grid1.CurrentCell.Value.ToString();
                saveFlag = 0;
                changeFlag[int.Parse(Grid1.CurrentRow.Cells[0].Value.ToString())] = 1;
            }
            if ((DateTime.Parse(Grid1.CurrentRow.Cells[5].Value.ToString()) - DateTime.Now).TotalDays <= 1 &&
                (DateTime.Parse(Grid1.CurrentRow.Cells[5].Value.ToString()) - DateTime.Now).TotalDays >= -1)
            {
                Grid1.CurrentRow.DefaultCellStyle.BackColor = Color.Red;
            }
            else
            {
                Grid1.CurrentRow.DefaultCellStyle.BackColor = Color.White;
            }
            Grid2Update();
        }

        private void Grid1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (Grid1.NewRowIndex == Grid1.CurrentRow.Index)
            {
                Grid1.CurrentRow.Cells[5].Value = DateTime.Now.AddDays(10).ToShortDateString();
            }
            if (Grid1.CurrentRow.Cells[0].Value != null) preNum = int.Parse(Grid1.CurrentRow.Cells[0].Value.ToString());
            else preNum = 0;
            if (Grid1.CurrentRow.Cells[2].Value == null) preMobile = "";
            else preMobile = Grid1.CurrentRow.Cells[2].Value.ToString();
            preDate = DateTime.Parse(Grid1.CurrentRow.Cells[5].Value.ToString());
        }

        private void save_result()
        {
            saveFlag = 1;
            string files = Application.StartupPath + "\\DataBase.mdb";
            string aConnectString = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=";
            aConnectString += files;
            sqlCon.ConnectionString = aConnectString;
            sqlCon.Open();
            sqlCmd.Connection = sqlCon;
            for (int ind = 0; ind < Grid1.RowCount; ind++)
            {
                if (Grid1.Rows[ind].Cells[0].Value == null) continue;
                int flagValue = changeFlag[int.Parse(Grid1.Rows[ind].Cells[0].Value.ToString())];
                if (flagValue == 1)
                {
                    string sqlstr = "select * from userData where 编号 = " + Grid1.Rows[ind].Cells[0].Value.ToString();
                    sqlAdp = new OleDbDataAdapter(sqlstr, sqlCon);
                    DataSet ds = new DataSet();
                    sqlAdp.Fill(ds, "userData");
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        sqlCmd.CommandText = UpdateStr(ind);
                        sqlCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        sqlCmd.CommandText = InsertStr(ind);
                        sqlCmd.ExecuteNonQuery();
                    }
                }
            }
            //记录需要删除的就删掉
            foreach (var item in changeFlag)
            {
                if (item.Value == 2)
                {
                    string sqlstr = "select * from userData where 编号 = " + item.Key;
                    sqlAdp = new OleDbDataAdapter(sqlstr, sqlCon);
                    DataSet ds = new DataSet();
                    sqlAdp.Fill(ds, "userData");
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        string delstr = "delete from userData where 编号 = " + item.Key;
                        sqlCmd.CommandText = delstr;
                        sqlCmd.ExecuteNonQuery();
                    }
                }
               
            }
            sqlCon.Close();
        }

        private string UpdateStr(int i)
        {
            string sqlstr = "update userData set 姓名 = '" + Grid1.Rows[i].Cells[1].Value + "', 电话 = '"
                    + Grid1.Rows[i].Cells[2].Value + "', 车型 = '" + Grid1.Rows[i].Cells[3].Value + "', 服务内容 = '"
                    + Grid1.Rows[i].Cells[4].Value + "', 服务时间 = '" + Grid1.Rows[i].Cells[5].Value +
                    "', 备注 = '" + Grid1.Rows[i].Cells[6].Value + "' where 编号 = " +
                    Grid1.Rows[i].Cells[0].Value;
            return sqlstr;
        }

        private string InsertStr(int i)
        {
            string sqlstr = "insert into userData(编号,姓名,电话,车型,服务内容,服务时间,备注)values(" + Grid1.Rows[i].Cells[0].Value + ",'"
                    + Grid1.Rows[i].Cells[1].Value + "','" + Grid1.Rows[i].Cells[2].Value + "','" + Grid1.Rows[i].Cells[3].Value + "','"
                    + Grid1.Rows[i].Cells[4].Value + "','" + Grid1.Rows[i].Cells[5].Value + "','" + Grid1.Rows[i].Cells[6].Value + "')";
            return sqlstr;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            save_result();
            MessageBox.Show("结果已保存！", "提示");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (startFlag == "false") //取消开机启动  
                {
                    string path = Application.StartupPath + "\\startupConfig.exe";
                    RegistryKey rk = Registry.LocalMachine;
                    RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                    rk2.DeleteValue("51chemi_DataManage", false);
                    rk2.Close();
                    rk.Close();

                }
                else //开机自启动
                {
                    string path = Application.StartupPath + "\\startupConfig.exe";
                    RegistryKey rk = Registry.LocalMachine;
                    RegistryKey rk2 = rk.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");
                    rk2.SetValue("51chemi_DataManage", path, RegistryValueKind.String);
                    rk2.Close();
                    rk.Close();
                }
            }
            catch { }
            if (File.Exists("startup.log")) File.Delete("startup.log");
            FileStream fw = new FileStream("startup.log", FileMode.OpenOrCreate);
            StreamWriter sw = new StreamWriter(fw);
            sw.WriteLine(startFlag);
            sw.Close();
            fw.Close();
            if (saveFlag == 0)
            {
                DialogResult dl = MessageBox.Show(this, "是否 保存结果(Save)？", "保存提示", MessageBoxButtons.YesNo);
                if (dl == DialogResult.Yes)
                {
                    save_result();
                    saveFlag = 1;
                } 
            }
            DialogResult dl2 = MessageBox.Show(this, "是否确定 退出(Exit)？", "退出提示", MessageBoxButtons.YesNo);
            if (dl2 == DialogResult.Yes) { }
            else
            {
                e.Cancel = true;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            saveFlag = 0;
            foreach (DataGridViewRow row in Grid1.SelectedRows)
            {
                changeFlag[int.Parse(row.Cells[0].Value.ToString())] = 2;
                Grid1.Rows.Remove(row);
            }
            sqlCon.Close();
            Grid2Update();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OleDbConnection con = new OleDbConnection();
            try
            {
                DialogResult dl = MessageBox.Show("导出数据将保存当前结果，是否导出？", "提示", MessageBoxButtons.YesNo);
                {
                    if(dl == DialogResult.No) return;
                }
                save_result();
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = ("Excel 文件(*.xls)|*.xls");//指定文件后缀名为Excel 文件。   
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    string filename = saveFile.FileName;
                    if (System.IO.File.Exists(filename))
                    {
                        System.IO.File.Delete(filename);//如果文件存在删除文件。   
                    }
                    //int index = filename.LastIndexOf("\\");//获取最后一个\的索引   
                    //filename = filename.Substring(index + 1);//获取excel名称(新建表的路径相对于SaveFileDialog的路径)   
                    //select * into 建立 新的表。   
                    //[[Excel 8.0;database= excel名].[sheet名] 如果是新建sheet表不能加$,如果向sheet里插入数据要加$.　   
                    //sheet最多存储65535条数据。   
                    string sql = "select top 65535 *  into   [Excel 8.0;database=" + filename + "].[用户信息] from userData order by 编号 asc";
                    con.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + Application.StartupPath + "\\DataBase.mdb";//将数据库放到debug目录下。   
                    OleDbCommand com = new OleDbCommand(sql, con);
                    con.Open();
                    com.ExecuteNonQuery();

                    MessageBox.Show("导出数据成功", "导出数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                con.Close();
            } 
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked) startFlag = "true";
            else startFlag = "false";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                if (pointIndex == 0) MessageBox.Show("已经是第一条了", "提示");
                else
                {
                    if (keyPosList.Count > 0)
                    {
                        pointIndex--;
                        Grid1.CurrentCell = Grid1[keyPosList[pointIndex].y, keyPosList[pointIndex].x];
                    }
                }
            }
            else
            {
                if (pointIndex2 == 0) MessageBox.Show("已经是第一条了", "提示");
                else
                {
                    if (keyPosList2.Count > 0)
                    {
                        pointIndex2--;
                        Grid2.CurrentCell = Grid2[keyPosList2[pointIndex2].y, keyPosList2[pointIndex2].x];
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                if (pointIndex == keyPosList.Count - 1) MessageBox.Show("已经是最后一条了", "提示");
                else
                {
                    if (keyPosList.Count > 0)
                    {
                        pointIndex++;
                        Grid1.CurrentCell = Grid1[keyPosList[pointIndex].y, keyPosList[pointIndex].x];
                    }
                }
            }
            else
            {
                if (pointIndex2 == keyPosList2.Count - 1) MessageBox.Show("已经是最后一条了", "提示");
                else
                {
                    if (keyPosList2.Count > 0)
                    {
                        pointIndex2++;
                        Grid2.CurrentCell = Grid2[keyPosList2[pointIndex2].y, keyPosList2[pointIndex2].x];
                    }
                }
            }   
        }

        private void Grid1_SortCompare(object sender, DataGridViewSortCompareEventArgs e)
        {
            //修改服务时间的比较方式为日期比较
            if (e.Column.Name == "服务时间")
            {
                e.SortResult = (Convert.ToDateTime(e.CellValue1) > Convert.ToDateTime(e.CellValue2)) ? 1 : (Convert.ToDateTime(e.CellValue1) < Convert.ToDateTime(e.CellValue2)) ? -1 : 0;
            }
            if (e.Column.Name == "编号")
            {
                e.SortResult = (Convert.ToInt32(e.CellValue1) > Convert.ToInt32(e.CellValue2)) ? 1 : (Convert.ToInt32(e.CellValue1) < Convert.ToInt32(e.CellValue2)) ? -1 : 0;
            }
            e.Handled = true;
        }

        private void Grid1_SelectionChanged(object sender, EventArgs e)
        {
            if (Grid1.CurrentCell.Style.BackColor == Color.Yellow) Grid1.DefaultCellStyle.SelectionBackColor = Color.Purple;
            else Grid1.DefaultCellStyle.SelectionBackColor = Color.Blue;
        }

        private void Grid1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            aPoint.x = Grid1.CurrentRow.Index;
            aPoint.y = Grid1.CurrentCell.ColumnIndex;
            if (keyPosList.Count > 0 && keyPosList.Contains(aPoint)) pointIndex = keyPosList.IndexOf(aPoint);
        }
    }
}
