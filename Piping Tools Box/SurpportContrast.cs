using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace Piping_Tools_Box
{
    public partial class SupportContrast : Form
    {
        public SupportContrast()
        {
            InitializeComponent();
        }
        //桌面路径
        string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        string supportTemplatePath = @"..\..\Template\SupportTemplate.xlsx";
        

        private void btnoldexcel_Click(object sender, EventArgs e)
        {
            init();
            //获取老表路径
            string oldPath = getExcelPath();
            tboldexcel.Text = oldPath;
            pgbRun.Value = 0;
        }        

        private void btnnewexcel_Click(object sender, EventArgs e)
        {
            //获取新表路径
            string newPath = getExcelPath();
            tbnewexcel.Text = newPath;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            
            //判断地址是否为空           
            if(tboldexcel.Text.Length==0|| tbnewexcel.Text.Length ==0)
            {
                MessageBox.Show("请在地址栏选择表格", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //显示运行状态错误
                runstate.Text = "运行错误";
                return;
            }
            //运行状态显示
            runstate.Text = "正在运行...";
            
            //打开老表
            Workbook oldworkbook = new Workbook(tboldexcel.Text);
            Worksheet oldworksheet = oldworkbook.Worksheets[0];
            Cells oldcells = oldworksheet.Cells;
            //打开新表
            Workbook newworkbook = new Workbook(tbnewexcel.Text);
            Worksheet newworksheet = newworkbook.Worksheets[0];
            Cells newcells = newworksheet.Cells;
            //引用模板表
            Workbook supportTemplateWorkbook = File.Exists(supportTemplatePath) ? new Workbook(supportTemplatePath) : new Workbook();
            Worksheet supportworksheet = supportTemplateWorkbook.Worksheets[0];
            Cells supportcells = supportworksheet.Cells;
            //新表最大数据行数
            int newRowCount = newcells.MaxDataRow;
            //老表最大数据行数
            int oldRowCount = oldcells.MaxDataRow;
            int j = 1;
            //创建字典，从而对比老表获取type
            Dictionary<string, string> oldDictionary = new Dictionary<string, string>();
            
            for(int m = 2; m <= oldRowCount; m++)
            {
                string oldsample = oldcells[m, 0].Value.ToString();
                if (!oldsample.Contains("SAMPLE"))
                {
                    string key = oldcells[m, 1].Value.ToString();//获取支架号为key
                    string value = oldcells[m, 6].Value.ToString();//获取type为value
                    oldDictionary.Add(key,value);//key value 放入字典
                }                
            }
           
            for (int i = 2; i <= newRowCount; i++)
            {
                string newsample = newcells[i, 0].Value.ToString();
                //过滤掉SAMPLE的支架号空值，
                if (!newsample.Contains("SAMPLE")){
                    supportcells[j, 0].PutValue(newcells[i, 1].Value);//支架号
                    supportcells[j, 1].PutValue(newcells[i, 6].Value);//type
                    //得到整理后的支架号
                    string cellValue = supportcells[j, 0].Value.ToString();
                    string typeValue = string.Empty;//存放字典获取的value
                    //设置颜色
                    //Style style = supportcells[0, 0].GetStyle();

                    //使用TryGetValue方法获取指定键对应的值,如果查到输出type，没有就为new支架号
                    if (oldDictionary.TryGetValue(cellValue,out typeValue))
                    {
                        supportcells[j, 2].PutValue(typeValue);
                    }
                    else
                    {
                        supportcells[j, 2].PutValue("NEW");
                        
                    }
                    //如果为不是新的支架号才对比，是新的支架号就不用管
                    if (supportcells[j,2].Value.ToString() != "NEW")
                    {
                        //对比老新type
                        string newtype = supportcells[j, 1].Value.ToString();
                        string oldtype = supportcells[j, 2].Value.ToString();

                        if (newtype == oldtype)
                        {
                            supportcells[j, 3].PutValue("OK");                            
                        }
                        else
                        {
                            string newbigtype = newtype.Substring(0, 3);
                            string oldbigtype = oldtype.Substring(0, 3);
                            if (newbigtype == oldbigtype)
                            {
                                supportcells[j, 3].PutValue("UPDATA");
                                //style.ForegroundColor = Color.Yellow;
                                //supportcells[j, 3].SetStyle(style);
                            }
                            else
                            {
                                supportcells[j, 3].PutValue("UPDATA DELETE NEW");
                                //style.ForegroundColor = Color.Red;
                                //supportcells[j, 3].SetStyle(style);
                            }
                        }
                    }
                    
                    //support表中行号
                    j++;
                    pgbRun.Minimum = 0;//进度条
                    pgbRun.Maximum = j;//最大值
                    pgbRun.Step = 1;
                    pgbRun.PerformStep();// 使用performstep方法按Step值递增
                }               
            }

            try
            {
                supportTemplateWorkbook.Save(strDesktopPath + @"\预制图变化清单.xlsx");
            }
            catch
            {
                MessageBox.Show("请关闭 预制图变化清单 表格", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
            //MessageBox.Show("工作已完成", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //完成运行状态
            runstate.Text = "已完成";            
        }

        //创建获取打开窗口的文件路径
        private string getExcelPath()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();//打开文件对话框
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "Excel文件(*.xls;*.xlsx;*xlsm)|*.xls;*.xlsx;*.xlsm";
            string path = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog.FileName;
            }
            return path;
        }
        //初始化
        private void init()
        {
            tboldexcel.Text = string.Empty;
            tbnewexcel.Text = string.Empty;
            pgbRun.Minimum = 0;//进度条
            runstate.Text = "等待运行";
        }

        private void btnOpenExcel_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(strDesktopPath + @"\预制图变化清单.xlsx");//打开桌面路径下的预制图变化清单.xlsx
        }
    }
}
