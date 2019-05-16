using Aspose.Cells;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Piping_Tools_Box
{
    public partial class PipMaterialCode : Form
    {
        public PipMaterialCode()
        {
            InitializeComponent();
        }
        //桌面路径
        string strDesktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        private void btnCheck_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();//打开文件对话框
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog.Filter = "Excel文件(*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string applyPath = openFileDialog.FileName;
                tbApply.Text = applyPath;
                if (applyPath.Contains("管子管件"))
                {
                    checkPipingApply(applyPath);
                }
                if (applyPath.Contains("螺栓垫片"))
                {
                    checkBoltApply(applyPath);
                }

            }
        }
        private void checkBoltApply(string applyPath)
        {
            Workbook workbook = new Workbook(applyPath);
            Worksheet descriptionSheet = workbook.Worksheets[1];
            Worksheet applySheet = workbook.Worksheets[2];
            Cells descriptionCells = descriptionSheet.Cells;
            Cells applyCells = applySheet.Cells;

            //设置一个格式，以便后期cell的引用红色                      
            Style mystyle = descriptionCells[1, 0].GetStyle();
            mystyle.ForegroundColor = Color.Red;

            //循环标准库，创建各个list
            //名称/name list
            //string[] name = { "STUD BOLTS WITH TWO NUTS", "GASKET" };
            ArrayList nameArrayList = new ArrayList();
            nameArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string nameValue = descriptionCells[k, 0].Value.ToString();
                nameArrayList.Add(nameValue);
            }
            //size list
            ArrayList sizeArrayList = new ArrayList();
            sizeArrayList.Add("-");
            for (int k = 1; k <= 39; k++)
            {
                string sizeValue = descriptionCells[k, 1].Value.ToString();
                sizeArrayList.Add(sizeValue);
            }
            //垫片压力等级/Pressure Class（LB)  list
            //string[] pressure = { "150LB", "300LB", "600LB", "900LB", "1500LB", "2500LB" };
            ArrayList pressureArrayList = new ArrayList();
            pressureArrayList.Add("-");
            for (int k = 1; k <= 39; k++)
            {
                string pressureValue = descriptionCells[k, 2].Value.ToString();
                pressureArrayList.Add(pressureValue);
            }
            //垫片端面/face list
            ArrayList faceArrayList = new ArrayList();
            faceArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string faceValue = descriptionCells[k, 3].Value.ToString();
                faceArrayList.Add(faceValue);
            }
            //垫片厚度/ Thickness(mm)  list
            ArrayList thicknessArrayList = new ArrayList();
            thicknessArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string thicknessValue = descriptionCells[k, 4].Value.ToString();
                thicknessArrayList.Add(thicknessValue);
            }
            //螺栓长度/ Length(mm) list
            ArrayList lengthArrayList = new ArrayList();
            lengthArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string lengthValue = descriptionCells[k, 5].Value.ToString();
                lengthArrayList.Add(lengthValue);
            }
            //第6+1列垫片类型/ Type list
            ArrayList typeArrayList = new ArrayList();
            typeArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string typeValue = descriptionCells[k, 6].Value.ToString();
                typeArrayList.Add(typeValue);
            }
            //第7+1列 材质描述/ Material Description list 自增超过跳出循环
            ArrayList descriptionArrayList = new ArrayList();
            descriptionArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string descriptionValue = descriptionCells[k, 7].Value.ToString();
                descriptionArrayList.Add(descriptionValue);
            }
            //第8+1列表面处理/ Surface Treatment
            ArrayList surfaceArrayList = new ArrayList();
            surfaceArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string surfaceValue = descriptionCells[k, 8].Value.ToString();
                surfaceArrayList.Add(surfaceValue);
            }
            //第9+1列颜色/ Colour list
            ArrayList colourArrayList = new ArrayList();
            colourArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string colourValue = descriptionCells[k, 9].Value.ToString();
                colourArrayList.Add(colourValue);
            }

            //第10+1列标准 /Standard list
            ArrayList standardArrayList = new ArrayList();
            standardArrayList.Add("-");
            for (int k = 1; k <= 19; k++)
            {
                string standardValue = descriptionCells[k, 10].Value.ToString();
                standardArrayList.Add(standardValue);
            }

            //开始循环申请表
            for (int i = 3; i <= applyCells.MaxDataRow; i++)
            {
                //第0+1列 名称name
                try
                {
                    Cell namecell = applyCells[i, 5];
                    string nameCellValue = namecell.Value.ToString();//第五列材料类别
                    if (nameCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        bool exists = nameArrayList.Contains(nameCellValue);
                        if (!exists)
                        {
                            namecell.PutValue("描述错误   " + nameCellValue);
                            namecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("名称/Name 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第1+1列尺寸size
                try
                {
                    Cell sizecell = applyCells[i, 6];
                    string sizeCellValue = sizecell.Value.ToString();
                    if (sizeCellValue != null)
                    {
                        bool exists = sizeArrayList.Contains(sizeCellValue);
                        if (!exists)
                        {
                            sizecell.PutValue("描述错误   " + sizeCellValue);
                            sizecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("尺寸/Size(Inch) 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第2+1列垫片压力等级
                try
                {
                    Cell pressurecell = applyCells[i, 7];
                    string pressureCellValue = pressurecell.Value.ToString();//第五列材料类别
                    if (pressureCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {

                        bool exists = pressureArrayList.Contains(pressureCellValue);
                        if (!exists)
                        {
                            pressurecell.PutValue("描述错误   " + pressureCellValue);
                            pressurecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("垫片压力等级/Pressure Class（LB) 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第3+1列垫片端面/ Face
                try
                {
                    Cell facecell = applyCells[i, 8];
                    string faceCellValue = facecell.Value.ToString();//第五列材料类别
                    if (faceCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] face = { "FF","RF","RTJ" };
                        bool exists = faceArrayList.Contains(faceCellValue);
                        if (!exists)
                        {
                            facecell.PutValue("描述错误   " + faceCellValue);
                            facecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("垫片端面/ Face 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第4+1列垫片厚度/ Thickness(mm)
                try
                {
                    Cell thicknesscell = applyCells[i, 9];
                    string thicknessCellValue = thicknesscell.Value.ToString();//第五列材料类别
                    if (thicknessCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] thickness = { "2", "3", "3.175" ,"4.5"};
                        bool exists = thicknessArrayList.Contains(thicknessCellValue);
                        if (!exists)
                        {
                            thicknesscell.PutValue("描述错误   " + thicknessCellValue);
                            thicknesscell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("垫片厚度/ Thickness(mm) 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

                //第5+1列 螺栓长度/ Length(mm)
                try
                {
                    Cell lengthcell = applyCells[i, 10];
                    string lengthCellValue = lengthcell.Value.ToString();
                    if (lengthCellValue != null)
                    {
                        bool exists = lengthArrayList.Contains(lengthCellValue);
                        if (!exists)
                        {
                            lengthcell.PutValue("描述错误   " + lengthCellValue);
                            lengthcell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("螺栓长度/ Length(mm) 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第6+1列垫片类型/ Type
                try
                {
                    Cell typecell = applyCells[i, 11];
                    string typeCellValue = typecell.Value.ToString();//第五列材料类别
                    if (typeCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] type = { "FLEXITALLIC TYPE \"CGI\"", "OVAL RING TYPE \"R\"", "FULL FACE", "FLAT RING RF" };
                        bool exists = typeArrayList.Contains(typeCellValue);
                        if (!exists)
                        {
                            typecell.PutValue("描述错误   " + typeCellValue);
                            typecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("垫片类型/ Type 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

                //第8+1列表面处理/ Surface Treatment
                try
                {
                    Cell surfacecell = applyCells[i, 13];
                    string surfaceCellValue = surfacecell.Value.ToString();//第五列材料类别
                    if (surfaceCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] surface = { "WHITFORD XYLAN 1014", "GALV." };
                        bool exists = surfaceArrayList.Contains(surfaceCellValue);
                        if (!exists)
                        {
                            surfacecell.PutValue("描述错误   " + surfaceCellValue);
                            surfacecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("表面处理/ Surface Treatment 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第9+1列颜色/ Colour
                try
                {
                    Cell colourcell = applyCells[i, 14];
                    string colourCellValue = colourcell.Value.ToString();//第五列材料类别
                    if (colourCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] colour = { "DEEP MED BLUE COLOR", "BLACK COLOR", "RED COLOR", "GREEN COLOR" };
                        bool exists = colourArrayList.Contains(colourCellValue);
                        if (!exists)
                        {
                            colourcell.PutValue("描述错误   " + colourCellValue);
                            colourcell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("颜色/ Colour 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第10+1列标准 /Standard 
                try
                {
                    Cell standardcell = applyCells[i, 15];
                    string standardCellValue = standardcell.Value.ToString();//第五列材料类别
                    if (standardCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        //string[] standard = { "ASME B18.2.1/B18.2.2", "ASME B16.21", "ASME B16.20"};
                        bool exists = standardArrayList.Contains(standardCellValue);
                        if (!exists)
                        {
                            standardcell.PutValue("描述错误   " + standardCellValue);
                            standardcell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("标准 /Standard  列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第7+1列 材质描述/ Material Description
                try
                {
                    Cell descriptioncell = applyCells[i, 12];
                    string descriptionCellValue = descriptioncell.Value.ToString();
                    if (descriptionCellValue != null)
                    {

                        bool exists = descriptionArrayList.Contains(descriptionCellValue);
                        if (!exists)
                        {
                            descriptioncell.PutValue("描述错误   " + descriptionCellValue);
                            descriptioncell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("材质描述/ Material Description 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            //for(int q = 3; q <= applyCells.MaxDataRow; q++)
            //{
            //    for(int w = 5; w <= applyCells.MaxDataColumn; w++)
            //    {
            //        string cellValue = applyCells[q, w].Value.ToString();
            //        if (cellValue.Contains("描述错误"))
            //        {
            //            Style styleRed = applyCells[q, w].GetStyle();
            //            styleRed.ForegroundColor = Color.Red;
            //            applyCells[q, w].SetStyle(styleRed);
            //        }
            //    }
            //}
            try
            {
                workbook.Save(strDesktopPath + @"\校验后-机管专业螺栓垫片编码申请表 .xlsx");
                MessageBox.Show("   螺栓垫片完成\n已在桌面保存文件", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("请关闭表格", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

        }

        private void checkPipingApply(string applyPath)
        {
            //创建工作表
            Workbook workbook = new Workbook(applyPath);
            Worksheet descriptionSheet = workbook.Worksheets[1];
            Worksheet applySheet = workbook.Worksheets[2];
            Cells descriptionCells = descriptionSheet.Cells;
            Cells applyCells = applySheet.Cells;
            //applyCells.ClearFormats(3, 0, applyCells.MaxDataRow, applyCells.MaxDataColumn);
            //设置一个格式，以便后期cell的引用红色,不知什么原因需要现在表格设置颜色也可以再去掉，才可以变颜色
            Style mystyle = descriptionCells[1, 0].GetStyle();
            mystyle.ForegroundColor = Color.Red;

            //材料描述的list
            ArrayList descriptionArrayList = new ArrayList();
            descriptionArrayList.Add("-");//应对与没有数据的cell，原表填写-
            for (int k = 1; k <= descriptionCells.MaxDataRow; k++)//循环标准库添加至各个list中
            {
                // MessageBox.Show(n.ToString(), "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                try
                {
                    string descriptionValue = descriptionCells[k, 4].Value.ToString();
                    descriptionArrayList.Add(descriptionValue);
                }
                catch
                {
                    continue;
                }

            }
            //catalogue list
            //string[] catalogue = { "CS", "LTCS", "GALV. CS", "SS", "DSS", "SDSS", "TI", "CUNI", "IT", "FRP", "COPPER", "UPVC", "-" };
            ArrayList catalogueArrayLsit = new ArrayList();
            catalogueArrayLsit.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string catalogueValue = descriptionCells[k, 0].Value.ToString();
                catalogueArrayLsit.Add(catalogueValue);
            }
            //将size基础数据放入ArrayList
            ArrayList sizeArrayList = new ArrayList();
            sizeArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string sizeValue = descriptionCells[k, 2].Value.ToString();
                sizeArrayList.Add(sizeValue);

            }
            //name list
            ArrayList nameArrayList = new ArrayList();
            nameArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string nameValue = descriptionCells[k, 1].Value.ToString();
                nameArrayList.Add(nameValue);
            }
            //pressure list
            //string[] pressure = { "150LB", "300LB", "600LB", "900LB", "1500LB", "2500LB", "3000LB", "6000LB", "9000LB", "16BAR", "-" };
            ArrayList pressureArrayList = new ArrayList();
            pressureArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string pressureValue = descriptionCells[k, 5].Value.ToString();
                pressureArrayList.Add(pressureValue);
            }
            //将壁厚基础数据放入ArrayList
            ArrayList thicknessArrayList = new ArrayList();
            thicknessArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string thicknessValue = descriptionCells[k, 6].Value.ToString();
                thicknessArrayList.Add(thicknessValue);
            }
            //连接形式connection list
            //string[] connection = { "PE", "BE", "BW", "SW", "FNPT", "MNPT", "FERRULE", "CB", "WN", "-" };
            ArrayList connectionArrayList = new ArrayList();
            connectionArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string connectionValue = descriptionCells[k, 8].Value.ToString();
                connectionArrayList.Add(connectionValue);
            }
            //法兰密封面形式
            //string[] flange = { "RF", "FF", "RTJ", "-" };
            ArrayList flangeArrayList = new ArrayList();
            flangeArrayList.Add("-");
            for (int k = 1; k <= 12; k++)
            {
                string flangeValue = descriptionCells[k, 10].Value.ToString();
                flangeArrayList.Add(flangeValue);
            }
            //弯曲半径
            //string[] bending = { "LR", "SR", "-" };
            ArrayList bendingArrayList = new ArrayList();
            bendingArrayList.Add("-");
            for (int k = 1; k <= 12; k++)
            {
                string bendingValue = descriptionCells[k, 11].Value.ToString();
                bendingArrayList.Add(bendingValue);
            }
            //表面处理
            //string[] surface = { "NACE", "GALV.", "ANNEALED", "FULLY ANNEALED (99 RB MAX.)", "-" };
            ArrayList surfaceArrayList = new ArrayList();
            surfaceArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string surfaceValue = descriptionCells[k, 12].Value.ToString();
                surfaceArrayList.Add(surfaceValue);
            }
            //标准
            ArrayList standardArrayList = new ArrayList();
            standardArrayList.Add("-");
            for (int k = 1; k <= 59; k++)
            {
                string standardValue = descriptionCells[k, 13].Value.ToString();
                standardArrayList.Add(standardValue);

            }

            //循环遍历所有apply表格数据对比
            for (int i = 3; i <= applyCells.MaxDataRow; i++)
            {
                //MessageBox.Show(applyCells.MaxDataRow.ToString(), "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                try
                {
                    Cell cataloguecell = applyCells[i, 5];
                    string catelogueCellValue = cataloguecell.Value.ToString();//第五列材料类别
                    if (catelogueCellValue != null)//防止格式的问题造成的Maxdatarow比实际值偏大
                    {
                        bool exists = catalogueArrayLsit.Contains(catelogueCellValue);
                        if (!exists)
                        {
                            //Style style = cataloguecell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //cataloguecell.SetStyle(style);
                            cataloguecell.PutValue("描述错误   " + catelogueCellValue);
                            cataloguecell.SetStyle(mystyle);
                            //MessageBox.Show("red", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("材料类别 /Material Catalogue 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第6+1列name
                try
                {
                    Cell namecell = applyCells[i, 6];
                    string nameCellValue = namecell.Value.ToString();
                    if (nameCellValue != null)
                    {

                        bool exists = nameArrayList.Contains(nameCellValue);
                        if (!exists)
                        {
                            //Style style = namecell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //namecell.SetStyle(style);
                            namecell.PutValue("描述错误   " + nameCellValue);

                            namecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("名称/Name 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

                //第7+1\8+1列main size\ 2nd size
                try
                {
                    Cell mainsizecell = applyCells[i, 7];
                    Cell secondsizecell = applyCells[i, 8];
                    string mainsizeCellValue = mainsizecell.Value.ToString();
                    string secondsizecellValue = secondsizecell.Value.ToString();

                    if (mainsizeCellValue != null)
                    {

                        bool existsmain = sizeArrayList.Contains(mainsizeCellValue);
                        if (!existsmain)
                        {
                            //主要尺寸设置
                            //Style styleMain = mainsizecell.GetStyle();
                            //styleMain.ForegroundColor = Color.Red;
                            //mainsizecell.SetStyle(styleMain);
                            mainsizecell.PutValue("描述错误   " + mainsizeCellValue);
                            mainsizecell.SetStyle(mystyle);
                        }
                    }
                    if (secondsizecellValue != null)
                    {
                        bool existssecond = sizeArrayList.Contains(secondsizecellValue);
                        if (!existssecond)
                        {
                            //次要尺寸设置
                            //Style styleSecond = secondsizecell.GetStyle();
                            //styleSecond.ForegroundColor = Color.Red;
                            //secondsizecell.SetStyle(styleSecond);
                            secondsizecell.PutValue("描述错误   " + secondsizecellValue);
                            secondsizecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("主尺寸或此尺寸 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }


                //第10+1列pressure class 压力等级
                try
                {
                    Cell pressurecell = applyCells[i, 10];
                    string pressureCellValue = pressurecell.Value.ToString();
                    if (pressureCellValue != null)
                    {
                        bool exists = pressureArrayList.Contains(pressureCellValue);

                        if (!exists)
                        {
                            //Style style = pressurecell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //pressurecell.SetStyle(style);
                            pressurecell.PutValue("描述错误   " + pressureCellValue);
                            pressurecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("压力等级 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //第11+1\12+1列主壁厚\次壁厚
                try
                {
                    Cell mainthicknesscell = applyCells[i, 11];
                    Cell secondthicknesscell = applyCells[i, 12];
                    string mainthicknessCellValue = mainthicknesscell.Value.ToString();
                    string secondthicknessCellValue = secondthicknesscell.Value.ToString();

                    if (mainthicknessCellValue != null)
                    {

                        bool mainexists = thicknessArrayList.Contains(mainthicknessCellValue);
                        if (!mainexists)
                        {
                            //主壁厚设置
                            //Style styleMain = mainthicknesscell.GetStyle();
                            //styleMain.ForegroundColor = Color.Red;
                            //mainthicknesscell.SetStyle(styleMain);  
                            mainthicknesscell.PutValue("描述错误   " + mainthicknessCellValue);
                            mainthicknesscell.SetStyle(mystyle);
                        }
                    }
                    if (secondthicknessCellValue != null)
                    {

                        bool secondexists = thicknessArrayList.Contains(secondthicknessCellValue);
                        if (!secondexists)
                        {
                            //主壁厚设置
                            //Style styleSecond = secondthicknesscell.GetStyle();
                            //styleSecond.ForegroundColor = Color.Red;
                            //secondthicknesscell.SetStyle(styleSecond);
                            secondthicknesscell.PutValue("描述错误   " + secondthicknessCellValue);
                            secondthicknesscell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("主壁厚或次壁厚 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //第13+1\14+1列主连接形式\次连接形式
                try
                {
                    Cell mainconnectioncell = applyCells[i, 13];
                    Cell secondconnectioncell = applyCells[i, 14];
                    string mainconnectionCellValue = mainconnectioncell.Value.ToString();
                    string secondconnectionCellValue = secondconnectioncell.Value.ToString();
                    if (mainconnectionCellValue != null)
                    {

                        bool mainexists = connectionArrayList.Contains(mainconnectionCellValue);
                        if (!mainexists)
                        {
                            //主连接形式设置
                            //Style styleMain = mainconnectioncell.GetStyle();
                            //styleMain.ForegroundColor = Color.Red;
                            //mainconnectioncell.SetStyle(styleMain);

                            mainconnectioncell.PutValue("描述错误   " + mainconnectionCellValue);
                            mainconnectioncell.SetStyle(mystyle);
                        }
                    }
                    if (secondconnectionCellValue != null)
                    {

                        bool secondexists = connectionArrayList.Contains(secondconnectionCellValue);
                        if (!secondexists)
                        {
                            //主连接形式设置
                            //Style styleSecond = mainconnectioncell.GetStyle();
                            //styleSecond.ForegroundColor = Color.Red;
                            //secondconnectioncell.SetStyle(styleSecond);
                            secondconnectioncell.PutValue("描述错误   " + secondconnectionCellValue);
                            secondconnectioncell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("主连接形式或次连接形式 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //15+1列法兰密封面形式
                try
                {
                    Cell flangecell = applyCells[i, 15];
                    string flangeCellValue = flangecell.Value.ToString();
                    if (flangeCellValue != null)
                    {

                        bool exists = flangeArrayList.Contains(flangeCellValue);
                        if (!exists)
                        {
                            //法兰设置
                            //Style style = flangecell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //flangecell.SetStyle(style);      
                            flangecell.PutValue("描述错误   " + flangeCellValue);
                            flangecell.SetStyle(mystyle);

                        }
                    }
                }
                catch
                {
                    MessageBox.Show("法兰密封面形式 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //16+1列弯曲半径
                try
                {
                    Cell bendingcell = applyCells[i, 16];
                    string bendingCellValue = bendingcell.Value.ToString();
                    if (bendingCellValue != null)
                    {
                        bool exists = bendingArrayList.Contains(bendingCellValue);
                        if (!exists)
                        {
                            //Style style = bendingcell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //bendingcell.SetStyle(style);
                            bendingcell.PutValue("描述错误   " + bendingCellValue);
                            bendingcell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("弯曲半径 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //17+1列表面处理
                try
                {
                    Cell surfacecell = applyCells[i, 17];
                    string surfaceCellValue = surfacecell.Value.ToString();
                    if (surfaceCellValue != null)
                    {
                        bool exists = surfaceArrayList.Contains(surfaceCellValue);
                        if (!exists)
                        {
                            //表面处理设置
                            //Style style = surfacecell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //surfacecell.SetStyle(style);
                            surfacecell.PutValue("描述错误   " + surfaceCellValue);
                            surfacecell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("表面处理 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                //第18+1列标准
                try
                {
                    Cell standardcell = applyCells[i, 18];
                    string standardCellValue = standardcell.Value.ToString();
                    if (standardCellValue != null)
                    {

                        bool exists = standardArrayList.Contains(standardCellValue);
                        if (!exists)
                        {
                            //Style style = standardcell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //standardcell.SetStyle(style);
                            standardcell.PutValue("描述错误   " + standardCellValue);
                            standardcell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("标准/Standard 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                //第9+1列Material Description  
                try
                {
                    Cell descriptioncell = applyCells[i, 9];
                    string descriptionCellValue = descriptioncell.Value.ToString();
                    if (descriptionCellValue != null)
                    {
                        bool exists = descriptionArrayList.Contains(descriptionCellValue);
                        if (!exists)
                        {
                            //Style style = descriptioncell.GetStyle();
                            //style.ForegroundColor = Color.Red;
                            //descriptioncell.SetStyle(style);
                            descriptioncell.PutValue("描述错误   " + descriptionCellValue);
                            descriptioncell.SetStyle(mystyle);
                        }
                    }
                }
                catch
                {
                    //MessageBox.Show("材料描述 列出现问题", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }

                //else
                //{
                //    MessageBox.Show("共检查到" + i + "行数据", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    break;
                //}

                //}
                //catch
                //{
                //    MessageBox.Show("文件格式有空行", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //}

            }

            try
            {
                workbook.Save(strDesktopPath + @"\校验后-机管专业管子管件编码申请表.xlsx");
                MessageBox.Show("   管子管件完成\n已在桌面保存文件", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("请关闭表格", "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (tbApply.Text.Contains("管子管件"))
            {
                System.Diagnostics.Process.Start(strDesktopPath + @"\校验后-机管专业管子管件编码申请表.xlsx");//打开桌面路径下的校验后的管子管件文件
            }
            if (tbApply.Text.Contains("螺栓垫片"))
            {
                System.Diagnostics.Process.Start(strDesktopPath + @"\校验后-机管专业螺栓垫片编码申请表 .xlsx");//打开桌面路径下的校验后的螺栓垫片文件
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("1.校验文件名需包含：管子管件或螺栓垫片，根据文件名进行区分\n2.如果报错，请删除表格中的空单元格或用-填充", "帮助提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
