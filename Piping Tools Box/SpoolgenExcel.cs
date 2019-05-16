using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;  
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;


namespace Piping_Tools_Box
{
    public partial class SpoolgenExcel : Form
    {        
        public SpoolgenExcel()
        {
            InitializeComponent();
        }
        //模板文件相对路径
        //string Template_File_Path = @"..\..\Template\Template.xlsx";
        string Template_TrackFile_Path = @"Template\Template_for_Piping_Material_Tracking_List.xlsx";//TrackingList模板相对路径
        string Template_WeldFile_Path = @"Template\Template_for_Piping_Welding_List.xlsx";//WeldingList模板相对路径
        string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);//桌面路径
        int i;
        int j;
        int trackFileCount = 0;
        int weldFileCount = 0;


        private void read_Click(object sender, EventArgs e)
        {
            initStart();//初始化界面
            //  打开 Excel 模板
            Workbook TrackTemplateWorkbook = File.Exists(Template_TrackFile_Path) ? new Workbook(Template_TrackFile_Path) : new Workbook();
            Workbook WeldTemplateWorkbook = File.Exists(Template_WeldFile_Path) ? new Workbook(Template_WeldFile_Path) : new Workbook();

            //  打开模板第一个sheet
            Worksheet trackTemplateSheet = TrackTemplateWorkbook.Worksheets[0];
            Worksheet weldTemplateSheet = WeldTemplateWorkbook.Worksheets[0];
            Cells trackCells = trackTemplateSheet.Cells;
            Cells weldCells = weldTemplateSheet.Cells;
            string folderPath=String.Empty;

            FolderBrowserDialog dialog = new FolderBrowserDialog();//打开文件夹对话框
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                folderPath = dialog.SelectedPath;
                tbBoxPath.Text = folderPath;
             
                DirectoryInfo root = new DirectoryInfo(folderPath);
                FileInfo[] files = root.GetFiles();              
                
                foreach (FileInfo file in files)
                {
                    string fileName = file.Name;
                    string fileFullName = file.FullName;
                   
                    if (fileName.Contains("Material"))
                    {
                        trackFileCount++;//文件个数
                        rtbTrack.AppendText(fileName+"\n");//将文件名添加至textbox中   
                        //Thread thread1 = new Thread(() => HandleMaterial(fileFullName, trackCells));
                        //thread1.Start();
                        HandleMaterial(fileFullName, trackCells);                      
                      
                    }
                    if (fileName.Contains("Welds"))
                    {
                        weldFileCount++;//文件个数
                        rtbWelds.AppendText(fileName + "\n");//将文件名添加至textbox中   
                        //Thread thread1 = new Thread(() => HandleMaterial(fileFullName, weldCells));
                        //thread1.Start();
                        HandleWelds(fileFullName, weldCells);

                    }

                    //读取的文件个数显示出来                    
                    tbTrackCount.Text = "共读取到"+" "+trackFileCount.ToString()+ " " + "个文件";
                    tbWeldsCount.Text = "共读取到" + " " + weldFileCount.ToString() + " " + "个文件";
                }
                try
                {
                    //保存文件
                    TrackTemplateWorkbook.Save(strDesktopPath+@"\Tracking List.xlsx");
                    WeldTemplateWorkbook.Save(strDesktopPath + @"\Welding List.xlsx");
                }
                catch
                {
                    MessageBox.Show("请关闭已打开的TrackingList或WeldingList表", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                tbTrackFinish.Text = "已完成";
                tbWeldFinish.Text = "已完成";                
            }
            
        }

        private void initStart()
        {
            i = 2;
            j = 2;
            tbTrackFinish.Text = "";
            tbWeldFinish.Text = "";
            trackFileCount= 0;
            weldFileCount = 0;
            pgbweld.Value=0;            
            rtbTrack.Text = string.Empty;
            rtbWelds.Text = string.Empty;
            tbBoxPath.Text = string.Empty;                
        }

        private void HandleWelds(string fileFullName, Cells weldCells)
        {
            string line = string.Empty;
            StreamReader reader = new StreamReader(fileFullName, Encoding.Default);
            while ((line = reader.ReadLine()) != null)
            {
                line=line.PadRight(1000);
                string col1 = line.Substring(0, 10).Trim();               
                    if (col1 == "COOECSIZE")
                {
                    continue;
                }
                string col2 = line.Substring(10, 40).Trim();
                string col3 = line.Substring(50, 25).Trim();
                string col4 = line.Substring(75, 25).Trim();
                string col5 = line.Substring(100, 25).Trim(); 
                string col6 = line.Substring(125, 25).Trim();
                string col7 = line.Substring(150, 25).Trim(); 
                string col8 = line.Substring(175, 25).Trim(); 
                string col9 = line.Substring(200, 25).Trim();
                string col10 = line.Substring(225, 10).Trim();
                string col11 = line.Substring(235, 15).Trim();
                string col12 = line.Substring(250, 25).Trim();
                string col13 = line.Substring(274, 50).Trim(); 
                string col14 = line.Substring(324, 50).Trim(); 
                string col15 = line.Substring(375, 25).Trim(); 
                string col16 = line.Substring(399, 25).Trim(); 
                string col17 = line.Substring(424, 25).Trim(); 
                string col18 = line.Substring(449, 25).Trim();
                string col19 = line.Substring(474, 25).Trim();
                string col20 = line.Substring(500, 145).Trim();
                string col21 = line.Substring(650, line.Length-650).Trim();
                string col22 = col6.Substring(0, col6.Length - 1);

                
                weldCells[j, 4 - 1].PutValue(col2);
                weldCells[j, 8 - 1].PutValue(col10);
                weldCells[j, 9 - 1].PutValue(col4);
                weldCells[j, 14 - 1].PutValue(col3);
                weldCells[j, 15 - 1].PutValue(col6);
                weldCells[j, 16 - 1].PutValue(col5);
                weldCells[j, 17 - 1].PutValue(col8);
                weldCells[j, 18 - 1].PutValue(col7);
                weldCells[j, 19 - 1].PutValue(col22);
                weldCells[j, 20 - 1].PutValue(col14);
                weldCells[j, 21 - 1].PutValue(col12);
                weldCells[j, 22 - 1].PutValue(col16);                
                weldCells[j, 23 - 1].PutValue(col20);
                weldCells[j, 24 - 1].PutValue(col15);
                weldCells[j, 25 - 1].PutValue(col13);
                weldCells[j, 26 - 1].PutValue(col17);
                weldCells[j, 27 - 1].PutValue(col21);

                j++;
                pgbweld.Minimum = 0;//进度条
                pgbweld.Maximum = j - 2;//最大值
                pgbweld.Step = 1;
                pgbweld.PerformStep();// 使用performstep方法按Step值递增
        }
        }

        private void HandleMaterial(string fileFullName,Cells trackCell)
        {
            string line = string.Empty;
            StreamReader reader = new StreamReader(fileFullName, Encoding.Default);
            
   
            while ((line = reader.ReadLine()) != null)
            {
                line=line.PadRight(1000);
                
                //cell和表格的对应关系是-1的关系
                string col1 = line.Substring(0, 10).Trim();
                if (col1 == "SPCCLASS")
                {
                    continue;
                }
                string col2 = line.Substring(10, 50).Trim(); 
                string col3 = line.Substring(60, 30).Trim(); 
                string col4 = line.Substring(90, 30).Trim();
                string col5 = line.Substring(120, 30).Trim();
                string col6 = line.Substring(150, 30).Trim(); 
                string col7 = line.Substring(180, 30).Trim();
                string col8 = line.Substring(210, 30).Trim();
                string col9 = line.Substring(240, 30).Trim();                
                string col10 = line.Substring(270, 30).Trim();
                string col11 = line.Substring(300, 30).Trim();
                string col12 = line.Substring(330, 29).Trim();
                string col13 = line.Substring(360, 30).Trim();
                string col14 = line.Substring(390, 29).Trim();
                string col15 = line.Substring(419, 30).Trim();
                string col16 = line.Substring(450, 30).Trim(); 
                string col17 = line.Substring(480, 30).Trim();
                string col18 = line.Substring(510, 30).Trim();
                string col19 = line.Substring(540, 30).Trim(); 
                string col20 = line.Substring(570, 30).Trim();
                string col21 = line.Substring(600, 30).Trim();
                string col22 = line.Substring(630, 29).Trim();
                string col23 = line.Substring(660, 30).Trim();
                string col24 = line.Substring(690, 29).Trim();
                string col25 = line.Substring(720, 30).Trim();
                string col26 = line.Substring(750, 30).Trim();
                string col27 = line.Substring(780, 30).Trim();
                string col28 = line.Substring(810, 140).Trim();
                string col29 = line.Substring(950, line.Length - 950).Trim();
                string[] descrition= col29.Split(',');

                trackCell[i, 15 - 1].PutValue(col2);
                trackCell[i, 16 - 1].PutValue(col4);
                trackCell[i, 26 - 1].PutValue(col6);
                trackCell[i, 25 - 1].PutValue(col6);
                if (col9.Contains("MM"))
                {
                    string[] unitArry9 = col9.Split(' ');
                    trackCell[i, 42 - 1].PutValue(unitArry9[1]);
                    
                    trackCell[i, 45 - 1].PutValue(unitArry9[0]);
                }
                else
                {
                    trackCell[i, 41 - 1].PutValue(col9);
                    trackCell[i, 42 - 1].PutValue("pcs");//单位
                }
                //string[] arry=cells[i, 14 - 1].Value.ToString().Split('x');//待定x的形式，x不存在问题
                //cells0[i, 31 - 1].PutValue(arry[0]);
                trackCell[i, 31 - 1].PutValue(col14);
                //cells0[i, 29 - 1].PutValue(arry[1]);
                trackCell[i, 30 - 1].PutValue(col15);
                trackCell[i, 57 - 1].PutValue(col16);
                trackCell[i, 10 - 1].PutValue(col19);
                trackCell[i, 18 - 1].PutValue(col25);
                trackCell[i, 20 - 1].PutValue(col26);
                trackCell[i, 35 - 1].PutValue(col27);
                trackCell[i, 55 - 1].PutValue(col28);
                trackCell[i, 72 - 1].PutValue(col29);

                i++;

                pgbtrack.Minimum = 0;//进度条
                pgbtrack.Maximum = i - 2;//最大值
                pgbtrack.Step = 1;
                pgbtrack.PerformStep();// 使用performstep方法按Step值递增
                
            }          

        }
        private void btnCheck_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(strDesktopPath + @"\Tracking List.xlsx");//打开桌面路径下的Tracking List

        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(strDesktopPath + @"\Welding List.xlsx");//打开桌面路径下的Weldinging List
        }

        private void tbBoxPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        
        //拖拽文件夹进入逻辑
        private void tbBoxPath_DragDrop(object sender, DragEventArgs e)
        {
            string FolderPath=((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            
            initStart();//初始化界面
            tbBoxPath.Text = FolderPath;
            //  打开 Excel 模板
            Workbook TrackTemplateWorkbook = File.Exists(Template_TrackFile_Path) ? new Workbook(Template_TrackFile_Path) : new Workbook();
            Workbook WeldTemplateWorkbook = File.Exists(Template_WeldFile_Path) ? new Workbook(Template_WeldFile_Path) : new Workbook();

            //  打开模板第一个sheet
            Worksheet trackTemplateSheet = TrackTemplateWorkbook.Worksheets[0];
            Worksheet weldTemplateSheet = WeldTemplateWorkbook.Worksheets[0];
            Cells trackCells = trackTemplateSheet.Cells;
            Cells weldCells = weldTemplateSheet.Cells;            
            
            DirectoryInfo root = new DirectoryInfo(FolderPath);
            FileInfo[] files = root.GetFiles();

            foreach (FileInfo file in files)
            {
                string fileName = file.Name;
                string fileFullName = file.FullName;

                if (fileName.Contains("Material"))
                {
                    trackFileCount++;//文件个数
                    rtbTrack.AppendText(fileName + "\n");//将文件名添加至textbox中   
                    //Thread thread1 = new Thread(() => HandleMaterial(fileFullName, trackCells));
                    //thread1.Start();
                    HandleMaterial(fileFullName, trackCells);

                }
                if (fileName.Contains("Welds"))
                {
                    weldFileCount++;//文件个数
                    rtbWelds.AppendText(fileName + "\n");//将文件名添加至textbox中   
                    //Thread thread1 = new Thread(() => HandleMaterial(fileFullName, weldCells));
                    //thread1.Start();
                    HandleWelds(fileFullName, weldCells);

                }

                //读取的文件个数显示出来                    
                tbTrackCount.Text = "共读取到" + " " + trackFileCount.ToString() + " " + "个文件";
                tbWeldsCount.Text = "共读取到" + " " + weldFileCount.ToString() + " " + "个文件";
            }
            try
            {
                //保存文件
                TrackTemplateWorkbook.Save(strDesktopPath + @"\Tracking List.xlsx");
                WeldTemplateWorkbook.Save(strDesktopPath + @"\Welding List.xlsx");
            }
            catch
            {
                MessageBox.Show("请关闭已打开的TrackingList或WeldingList表", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            tbTrackFinish.Text = "已完成";
            tbWeldFinish.Text = "已完成";
        }        
    }
}
