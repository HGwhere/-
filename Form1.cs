/*
 * 
 * 
 * 2024.5.1
 * 完成总体框架，至少能让程序跑起来了
 * 
 * 
 */
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Runtime.CompilerServices;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using OfficeOpenXml.DataValidation;
using System.Globalization;
using System.Data.Common;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Presentation;
using System.IO.Packaging;
using static ClosedXML.Excel.XLPredefinedFormat;



namespace WindowsFormsApp2
{

    public partial class Form1 : Form
    {
        string FilePath, Gread, FilePath_Grade;
        List<string> column1Data = new List<string>();
        List<string> column2Data = new List<string>();
        List<int> column3Data = new List<int>();

        public Form1()
        {
            InitializeComponent();
            textBox1.TextChanged += new EventHandler(textBox1_TextChanged);
            comboBox1.SelectedIndexChanged += new EventHandler(comboBox1_SelectedIndexChanged);
            textBox1.TextChanged += textBox1_TextChanged;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;//非商用许可
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void Form1_Resize(object sender, EventArgs e)
        {

        }
        private void button5_Click(object sender, EventArgs e) //第一步下一页
        {
            // 加载Excel文件  
            // 打开刚刚打开的文件，收集表的收集内容肯定不是按照学号排序的，所以现在开始排序
            using (var package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

                // 读取所有行的数据，并创建一个包含所有行的列表  
                var rows = new List<Dictionary<int, object>>();
                int startRow = 2; // 假设第一行是标题行，从第二行开始读取数据  
                int endRow = worksheet.Dimension.End.Row;

                for (int row = startRow; row <= endRow; row++)
                {
                    var rowData = new Dictionary<int, object>();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowData[col] = worksheet.Cells[row, col].Value;
                    }
                    rows.Add(rowData);
                }

                // 根据C列的数字大小进行排序  
                rows = rows
                    .OrderBy(r =>
                    {
                        if (r.ContainsKey(2) && r[3] != null) // 确保C列有值且不为null  
                        {
                            return Convert.ToDouble(r[3]); // 假设C列是数字，转换为double进行排序  
                        }
                        return double.MaxValue;  // 如果C列为空或无法转换，则放在最后  
                    })
                    .ToList();

                // 清空除标题行外的所有行数据  
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        worksheet.Cells[row, col].Value = null;
                    }
                }

                // 将排序后的数据写回Excel中  
                int rowToWrite = startRow - 1; // 标题行不需要重新写，所以从第一行数据开始写  
                foreach (var rowData in rows)
                {
                    rowToWrite++;
                    foreach (var kvp in rowData)
                    {
                        worksheet.Cells[rowToWrite, kvp.Key].Value = kvp.Value;
                    }
                }

                // 保存修改后的Excel文件  
                package.SaveAs(new FileInfo(FilePath)); // 如果需要保存到新的文件，可以传递新的文件路径给SaveAs方法  
            }
            //把刚刚排完序的工作表复制一份，专门用来处理备注所需要的文本
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                // 获取要复制的工作表  
                ExcelWorksheet worksheetToCopy = package.Workbook.Worksheets[0];

                // 添加一个新工作表  
                ExcelWorksheet newWorksheet = package.Workbook.Worksheets.Add("文本生成");

                // 复制所有单元格的值  
                for (int row = 1; row <= worksheetToCopy.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheetToCopy.Dimension.End.Column; col++)
                    {
                        // 获取旧工作表的单元格  
                        ExcelRangeBase oldCell = worksheetToCopy.Cells[row, col];

                        // 获取新工作表的对应单元格，并复制值  
                        ExcelRangeBase newCell = newWorksheet.Cells[row, col];
                        newCell.Value = oldCell.Value;

                        // 如果需要复制格式，可以添加如下代码：  
                        // newCell.Style.CopyFrom(oldCell.Style);  

                        // 如果需要复制公式，并且旧单元格有公式，可以添加如下代码：  
                        // if (oldCell.HasFormula)  
                        // {  
                        //     newCell.Formula = oldCell.Formula;  
                        // }  
                    }
                }

                // 保存更改到新的Excel文件或者覆盖原始文件  
                new FileInfo(FilePath);
                package.SaveAs(FilePath);
            }
            tabControl1.SelectedIndex = 1;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void button1_Click_1(object sender, EventArgs e)   //第一页的导入文件按钮
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "选择导入文件";
            //fdlg.InitialDirectory = @"c:\";   //@是取消转义字符的意思  
            fdlg.Filter = "Excel文档(*.xlsx;*.xls)|*.xlsx;*.xls";
            /* 
             * FilterIndex 属性用于选择了何种文件类型,缺省设置为0,系统取Filter属性设置第一项 
             * ,相当于FilterIndex 属性设置为1.如果你编了3个文件类型，当FilterIndex ＝2时是指第2个. 
             */
            fdlg.FilterIndex = 2;
            /* 
             *如果值为false，那么下一次选择文件的初始目录是上一次你选择的那个目录， 
             *不固定；如果值为true，每次打开这个对话框初始目录不随你的选择而改变，是固定的   
             */
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                //label3.Text = System.IO.Path.GetFileNameWithoutExtension(fdlg.FileName);//没有扩展名
                label3.Text = System.IO.Path.GetFileName(fdlg.FileName);//有扩展名
                FilePath = System.IO.Path.GetFullPath(fdlg.FileName); //获取文件路径
                //MessageBox.Show(FilePath);
            }
            //来源：https://blog.csdn.net/s_156/article/details/134156275


        }

        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void label1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)   //记录
        {
            // 清空DataGridView中现有的行（如果需要的话）  
            //dataGridView1.Rows.Clear();
            string[] lines = textBox1.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            string[] lines1 = textBox2.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (lines.Length == lines1.Length)
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    string trimmedLine = lines[i].Trim(); // 去除首尾空格  
                    string trimmedLine1 = lines1[i].Trim(); // 去除首尾空格  

                    if (!string.IsNullOrEmpty(trimmedLine) && !string.IsNullOrEmpty(trimmedLine1))
                    {
                        int rowIndex = dataGridView1.Rows.Add(); // 在最后一行添加新行并获取其索引  

                        // 设置新行各列的值  
                        dataGridView1.Rows[rowIndex].Cells[0].Value = trimmedLine; // 第一列的值  
                        dataGridView1.Rows[rowIndex].Cells[1].Value = comboBox1.Text; // 第二列的值  
                        dataGridView1.Rows[rowIndex].Cells[2].Value = trimmedLine1; // 第三列的值  
                    }
                }
            }
            else
            {
                // 处理lines和lines1长度不同的情况  
                MessageBox.Show("活动名称与分值两栏行数不同，无法录入。");
            }
            textBox1.Text = string.Empty;
            textBox2.Text = string.Empty;
        }

        private void button3_Click(object sender, EventArgs e)  //重置警告
        {
            //来源：https://blog.csdn.net/lucgh/article/details/130161414
            if (MessageBox.Show("确定重置表格？您将重新填写刚刚所录入的全部数据!", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
            {
                dataGridView1.Rows.Clear();
            }
        }

        private void button4_Click(object sender, EventArgs e)   //保存数据并下一步
        {
            /*
             * 
             * 接下来是：DataGridView>>遍历DataGridView的匹配项>>替换
             * 
             */
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // 检查这一行不是新行（没有数据的行）  
                if (!row.IsNewRow)
                {
                    // 将第一列和第二列的数据添加到对应的列表中  
                    column1Data.Add(row.Cells[0].Value as string);
                    column2Data.Add(row.Cells[1].Value as string);

                    // 将第三列的数据转换为int并添加到对应的列表中  
                    if (row.Cells[2].Value != null && int.TryParse(row.Cells[2].Value.ToString(), out int intValue))
                    {
                        column3Data.Add(intValue);
                    }
                    else
                    {
                        // 如果转换失败，可以添加错误处理或者跳过这一行  
                        // 例如：throw new InvalidOperationException("无法将第三列的数据转换为整数。");  
                        // 或者：continue; // 跳过当前行  
                    }
                }
            }
            // 假设column1Data, column2Data, column3Data已经被填充了数据  

            // 检查列表长度是否一致  
            if (column1Data.Count == column2Data.Count && column2Data.Count == column3Data.Count)
            {
                // 遍历列表并打印出映射关系  
                for (int i = 0; i < column1Data.Count; i++)
                {
                    string rowData = $"行 {i + 1}: 列1 = {column1Data[i]}, 列2 = {column2Data[i]}, 列3 = {column3Data[i]}";
                    System.Diagnostics.Debug.WriteLine(rowData); // 在控制台输出  
                                                                 // 或者，如果你想在WinForms中显示，可以使用如下方式：  
                                                                 //MessageBox.Show(rowData); // 假设你有一个名为listBox1的ListBox控件  
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("映射关系不一致，无法查看。");
            }
            //MessageBox.Show("数据已保存！");//完成测试
            tabControl1.SelectedIndex = 2;
        }
        //************************************************************************
        //
        //
        // 
        //第三页
        //
        //
        //
        //************************************************************************
        private void button8_Click(object sender, EventArgs e)   //生成成绩模板
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                // 创建一个新的工作表，并将其添加到Excel文件中  
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("成绩");

                // 在第一行第一列添加标题  

                //// 添加一些数据行  
                //worksheet.Cells[2, 1].Value = 1;
                //worksheet.Cells[2, 2].Value = "John Doe";
                //worksheet.Cells[2, 3].Value = 30;

                //worksheet.Cells[3, 1].Value = 2;
                //worksheet.Cells[3, 2].Value = "Jane Smith";
                //worksheet.Cells[3, 3].Value = 25;

                //// 设置列宽（可选）  
                //worksheet.Column(1).Width = 10;
                //worksheet.Column(2).Width = 25;
                //worksheet.Column(3).Width = 15;

                // 将Excel文件保存到磁盘上  
                //FileInfo excelFile = new FileInfo("成绩.xlsx");
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx"; // 设置文件过滤器为Excel文件  
                saveFileDialog1.FileName = "成绩.xlsx"; // 设置默认文件名  
                saveFileDialog1.Title = "请选择保存成绩文件的地址"; // 设置对话框标题  
                // 显示保存文件对话框  
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = saveFileDialog1.FileName;

                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        // 创建一个新的工作表  
                        ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets.Add("成绩");

                        // ... 在这里添加你的行和列数据到worksheet ...  
                        worksheet1.Cells[2, 1].Value = "姓名";
                        worksheet1.Cells[2, 2].Value = "学号";
                        worksheet1.Cells[2, 3].Value = "各科对应学分";
                        worksheet1.Cells[1, 4].Value = "科目1名称";
                        worksheet1.Cells[1, 5].Value = "科目2名称";
                        worksheet1.Cells[1, 6].Value = "以此类推，插入各个科目名称，最大支持15个科目（R列）";
                        // 保存ExcelPackage对象到文件  
                        FileInfo excelFileInfo = new FileInfo(filePath);
                        excelPackage.SaveAs(excelFileInfo);

                        // 无需显式调用Dispose，因为using语句会自动处理  
                    }

                    MessageBox.Show("Excel文件已保存到指定位置。");
                }
                else
                {
                    MessageBox.Show("未选择保存地址。");
                }
                //package.SaveAs(excelFile);
            }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)    //成绩文件选取
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "选择导入文件";
            //fdlg.InitialDirectory = @"c:\";   //@是取消转义字符的意思  
            fdlg.Filter = "Excel文档(*.xlsx;*.xls)|*.xlsx;*.xls";
            /* 
             * FilterIndex 属性用于选择了何种文件类型,缺省设置为0,系统取Filter属性设置第一项 
             * ,相当于FilterIndex 属性设置为1.如果你编了3个文件类型，当FilterIndex ＝2时是指第2个. 
             */
            fdlg.FilterIndex = 2;
            /* 
             *如果值为false，那么下一次选择文件的初始目录是上一次你选择的那个目录， 
             *不固定；如果值为true，每次打开这个对话框初始目录不随你的选择而改变，是固定的   
             */
            fdlg.RestoreDirectory = false;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                //label3.Text = System.IO.Path.GetFileNameWithoutExtension(fdlg.FileName);//没有扩展名
                label10.Text = System.IO.Path.GetFileName(fdlg.FileName);//有扩展名
                FilePath_Grade = System.IO.Path.GetFullPath(fdlg.FileName); //获取文件路径
                //MessageBox.Show(FilePath_Grade);
            }
            //来源：https://blog.csdn.net/s_156/article/details/134156275
        }

        private void button6_Click(object sender, EventArgs e)   //完成并开始计算  P3
        {
            // 加载源工作簿和目标工作簿  
            using (ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(FilePath_Grade)))
            {
                ExcelWorksheet sourceWorksheet = sourcePackage.Workbook.Worksheets["成绩"];
                if (sourceWorksheet == null)
                {
                    MessageBox.Show("没找到成绩");
                    return;
                }
                else
                {
                    using (ExcelPackage targetPackage = new ExcelPackage(new FileInfo(FilePath)))
                    {
                        // 将源工作表复制到目标工作簿中  
                        ExcelWorksheet copiedWorksheet = targetPackage.Workbook.Worksheets.Add(sourceWorksheet.Name, sourceWorksheet);

                        // 保存目标工作簿  
                        FileInfo targetFileInfo = new FileInfo(FilePath);
                        targetPackage.SaveAs(targetFileInfo);
                    }
                }
            }
            using (ExcelPackage package = new ExcelPackage(FilePath))
            {
                // 获取第一个工作表（通常是默认的工作表）  
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];

                int maxColWithContent = worksheet.Dimension.End.Column; // 获取成绩工作表的最大列数  
                System.Diagnostics.Debug.WriteLine($"成绩的最大一列是：{maxColWithContent}");

                //获取最大空行数
                int maxRowWithContent = worksheet.Dimension.End.Row; // 获取成绩工作表的最大行数  
                System.Diagnostics.Debug.WriteLine("成绩的最大非空行数是: " + maxRowWithContent);

                // 确定求和范围（例如，第5行的A列到最右边的非空单元格）  
                ExcelRangeBase range = worksheet.Cells["D2:Z2"];   //最大z列

                double sum = 0;

                // 遍历范围并手动计算总和  
                foreach (var cell in range)
                {
                    if (cell.Value != null && (cell.Value is double || cell.Value is int || cell.Value is decimal || (cell.Value is string s && double.TryParse(s, out double result))))
                    {
                        sum += Convert.ToDouble(cell.Value);
                    }
                }
                // 将计算结果放在G5单元格中  
                worksheet.Cells["C2"].Value = sum;

                // 保存对Excel文件的更改  
                package.Save();

                int row_now = 3;  //行
                while (row_now <= maxRowWithContent)  //如果现在的行还没到工作表的最后一个非空行，那么
                {
                    for (int col = 4; col <= maxColWithContent; col++)   //检查当前列
                    {
                        // 获取当前列的单元格  
                        ExcelRangeBase cell = worksheet.Cells[row_now, col];

                        // 检查单元格是否为空,如果为空，那么换行检测
                        if (cell.Value != null && double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double cellValue))
                        {
                            System.Diagnostics.Debug.WriteLine("当前在第" + row_now + "行，第" + col + "列");
                            if (cellValue < 60)   //如果分数小于60
                            {
                                System.Diagnostics.Debug.WriteLine("有小于60的");
                                if (worksheet.Cells[row_now, col].Value == null)  //如果挂科数单元格内为Null
                                {
                                    worksheet.Cells[row_now, 3].Value = 2;  //那么为2，因为Null不能直接加减
                                    System.Diagnostics.Debug.WriteLine("已添加数字1");
                                }
                                else
                                {
                                    object guakeshu = worksheet.Cells[row_now, 3].Value;
                                    int guake = Convert.ToInt16(guakeshu);
                                    guake = guake + 2;  //要是已经有一门挂科的，那么在1的基础上直接加1就行
                                    worksheet.Cells[row_now, 3].Value = guake;
                                }
                            }

                        }
                        if(col == maxColWithContent)
                        {
                            object guakeshujiance = worksheet.Cells[row_now, 3].Value;
                            int guakejiance = Convert.ToInt16(guakeshujiance);
                            if (guakejiance == 0)
                            {
                                worksheet.Cells[row_now, 3].Value = 0;
                            }
                        }
                        package.SaveAs(FilePath);
                    }
                    row_now++;
                }
                // 创建乘积之和。  

                worksheet.Cells[2, maxColWithContent + 1].Value = "乘积之和";
                worksheet.Cells[2, maxColWithContent + 2].Value = "减去挂科";
                package.SaveAs(FilePath);
                // 初始化求和变量  
                int rowNumber = 3; // 要读取的行号
                int startColumn = 4;
                int end_row = 3;
                int end_col = 4;
                // 创建一个double类型的列表来存储数据  

                List<double> values_xuefen = new List<double>();
                for (int col = startColumn; col <= maxColWithContent; col++)
                {
                    ExcelRange cell = worksheet.Cells[2, startColumn];
                    // 尝试将单元格的值转换为double  
                    if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double value_xuefen))
                    {
                        values_xuefen.Add(value_xuefen);
                        startColumn++;
                    }
                    else
                    {
                        values_xuefen.Add(0); // 添加默认值0  
                        startColumn++;
                    }
                }
                foreach (double value in values_xuefen)
                {
                    Console.WriteLine($"学分{value}");
                }
                Console.WriteLine("好了");
                startColumn = 4;
                //求得成绩与学分相乘之和
                List<double> values_chengji = new List<double>();
                List<double> sum_final = new List<double>();
                List<double> sum_final1 = new List<double>();
                List<double> sum_final2 = new List<double>();
                double final,final1,final2;
                final = 0;
                final1 = 0;
                final2 = 0;
                while (end_row <= maxRowWithContent)
                {
                    if (end_col <= maxColWithContent)
                    {
                        //成绩列  存入列表
                        ExcelRange cell_chengji = worksheet.Cells[end_row, end_col];
                        // 尝试将单元格的值转换为double  
                        if (cell_chengji.Value != null && double.TryParse(cell_chengji.Value.ToString(), out double value_chengji))
                        {
                            values_chengji.Add(value_chengji);
                            end_col++;
                        }
                        else
                        {
                            values_chengji.Add(0); // 添加默认值0  
                            end_col++;
                        }
                    }
                    Console.WriteLine("卡住了");
                    if (end_col > maxColWithContent)  //已经完全存到列表里了
                    {
                        //开始第一步计算
                        for (int i = 0; i < maxColWithContent - 3; i++)  //计算科目成绩与学分的乘积之和
                        {

                            final = final + (values_chengji[i] * values_xuefen[i]);
                            //sum_final.Add(final);
                            Console.WriteLine($"{values_chengji[i]}({values_chengji.Count}) * {values_xuefen[i]}({values_xuefen.Count})  ------ 结果为: {final}");
                        }
                        worksheet.Cells[end_row, maxColWithContent + 1].Value = final;   //先存入最右边的空列中
                        //计算完成后将该列的乘积之和减去挂科数（已经提前*2）
                        ExcelRange cell_sumchengji = worksheet.Cells[end_row, maxColWithContent + 1];    //获取当前行的乘积之和（获取刚刚存入空列的值）
                        ExcelRange cell_guakechengji = worksheet.Cells[end_row, 3];    //获取当前行的挂科值
                        ExcelRange cell_zongxuefen = worksheet.Cells["C2"];    //获取总学分
                        // 尝试将单元格的值转换为double  
                        double.TryParse(cell_sumchengji.Value.ToString(), out double value_sumchengji);
                        double.TryParse(cell_guakechengji.Value.ToString(), out double value_guakechengji);
                        double.TryParse(cell_zongxuefen.Value.ToString(), out double value_zongxuefen);

                        value_sumchengji = (value_sumchengji - value_guakechengji) / value_zongxuefen;
                        worksheet.Cells[end_row, maxColWithContent + 2].Value = value_sumchengji;


                        final = 0;
                        end_col = 4;
                        values_chengji.Clear();
                        end_row++;
                        package.SaveAs(FilePath);
                    }
                }
                package.SaveAs(FilePath);
            }
            /*
             * 
             * 完成第一部分
             * 开始将加分项与分数替换
             * 
             * 
             * 
             */
            using (ExcelPackage excelPackage = new ExcelPackage(FilePath))
            {
                // 确保工作簿中至少有两个工作表  
                if (excelPackage.Workbook.Worksheets.Count >= 2)
                {
                    if (column1Data.Count == column3Data.Count)
                    {
                        // 加载Excel文件  
                        using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
                        {
                            // 获取第一个工作表  
                            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                            // 遍历工作表中的所有单元格  
                            foreach (var cell in worksheet.Cells)
                            {
                                // 检查单元格是否为空  
                                if (!string.IsNullOrEmpty(cell.Value?.ToString()))
                                {
                                    // 遍历要搜索和替换的字符串列表  
                                    for (int i = 0; i < column1Data.Count; i++)
                                    {
                                        // 检查单元格内容是否包含要搜索的字符串  
                                        if (cell.Value.ToString().Contains(column1Data[i]))
                                        {
                                            // 将整数值转换为字符串  
                                            string replacement = column3Data[i].ToString();

                                            // 替换单元格内容中与搜索字符串匹配的部分  
                                            cell.Value = cell.Value.ToString().Replace(column1Data[i], replacement);
                                            //break; // 如果添加这行，会导致每一个单元格只有第一个活动会被识别与替换
                                        }
                                    }
                                }
                                else
                                {
                                    cell.Value = "0";
                                }
                            }
                            // 保存更改到Excel文件  
                            package.SaveAs(FilePath);
                            System.Diagnostics.Debug.WriteLine("替换操作已完成。");
                        }
                    }
                }

            }
            /*
             * 
             * 
             * 第二部分：对单元格内的分数进行计算
             * 
             * 
             */
            if (!File.Exists(FilePath))
            {
                System.Diagnostics.Debug.WriteLine("文件不存在，请检查路径是否正确。");
                return;
            }
            // 使用 ExcelPackage 打开工作簿  
            using (ExcelPackage package = new ExcelPackage(FilePath))
            {
                // 检查工作簿是否包含至少一个工作表  
                if (package.Workbook.Worksheets.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("工作簿是空的，将创建一个新的工作表。");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("工作簿已包含工作表，将添加一个新的工作表。");
                }

                // 创建一个新的工作表  
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("综测总分");


                // 计算起始和结束单元格的行号和列号  
                //标题
                ExcelRangeBase Title_startCell = worksheet.Cells["A1"];
                ExcelRangeBase Title_endCell = worksheet.Cells["O1"];
                int Title_startRow = Title_startCell.Start.Row;
                int Title_startCol = Title_startCell.Start.Column;
                int Title_endRow = Title_endCell.End.Row;
                int Title_endCol = Title_endCell.End.Column;
                //序号
                ExcelRangeBase Num_startCell = worksheet.Cells["A2"];
                ExcelRangeBase Num_endCell = worksheet.Cells["A3"];
                int Num_startRow = Num_startCell.Start.Row;
                int Num_startCol = Num_startCell.Start.Column;
                int Num_endRow = Num_endCell.End.Row;
                int Num_endCol = Num_endCell.End.Column;
                //姓名
                ExcelRangeBase Name_startCell = worksheet.Cells["B2"];
                ExcelRangeBase Name_endCell = worksheet.Cells["B3"];
                int Name_startRow = Name_startCell.Start.Row;
                int Name_startCol = Name_startCell.Start.Column;
                int Name_endRow = Name_endCell.End.Row;
                int Name_endCol = Name_endCell.End.Column;
                //学号
                ExcelRangeBase StuID_startCell = worksheet.Cells["C2"];
                ExcelRangeBase StuID_endCell = worksheet.Cells["C3"];
                int StuID_startRow = StuID_startCell.Start.Row;
                int StuID_startCol = StuID_startCell.Start.Column;
                int StuID_endRow = StuID_endCell.End.Row;
                int StuID_endCol = StuID_endCell.End.Column;
                //专业班级
                ExcelRangeBase Class_startCell = worksheet.Cells["D2"];
                ExcelRangeBase Class_endCell = worksheet.Cells["D3"];
                int Class_startRow = Class_startCell.Start.Row;
                int Class_startCol = Class_startCell.Start.Column;
                int Class_endRow = Class_endCell.End.Row;
                int Class_endCol = Class_endCell.End.Column;
                //总分
                ExcelRangeBase Gread_startCell = worksheet.Cells["M2"];
                ExcelRangeBase Gread_endCell = worksheet.Cells["M3"];     //M3!
                int Gread_startRow = Gread_startCell.Start.Row;
                int Gread_startCol = Gread_startCell.Start.Column;
                int Gread_endRow = Gread_endCell.End.Row;
                int Gread_endCol = Gread_endCell.End.Column;
                //总分排名
                ExcelRangeBase GreadRank_startCell = worksheet.Cells["N2"];
                ExcelRangeBase GreadRank_endCell = worksheet.Cells["N3"];
                int GreadRank_startRow = GreadRank_startCell.Start.Row;
                int GreadRank_startCol = GreadRank_startCell.Start.Column;
                int GreadRank_endRow = GreadRank_endCell.End.Row;
                int GreadRank_endCol = GreadRank_endCell.End.Column;
                //思想品德
                ExcelRangeBase sixiang_startCell = worksheet.Cells["E2"];
                ExcelRangeBase sixiang_endCell = worksheet.Cells["F2"];
                int sixiang_startRow = sixiang_startCell.Start.Row;
                int sixiang_startCol = sixiang_startCell.Start.Column;
                int sixiang_endRow = sixiang_endCell.End.Row;
                int sixiang_endCol = sixiang_endCell.End.Column;
                //专业学习
                ExcelRangeBase zhuanye_startCell = worksheet.Cells["G2"];
                ExcelRangeBase zhuanye_endCell = worksheet.Cells["H2"];
                int zhuanye_startRow = zhuanye_startCell.Start.Row;
                int zhuanye_startCol = zhuanye_startCell.Start.Column;
                int zhuanye_endRow = zhuanye_endCell.End.Row;
                int zhuanye_endCol = zhuanye_endCell.End.Column;
                //文体活动
                ExcelRangeBase wenti_startCell = worksheet.Cells["I2"];
                ExcelRangeBase wenti_endCell = worksheet.Cells["J2"];
                int wenti_startRow = wenti_startCell.Start.Row;
                int wenti_startCol = wenti_startCell.Start.Column;
                int wenti_endRow = wenti_endCell.End.Row;
                int wenti_endCol = wenti_endCell.End.Column;
                //社会服务
                ExcelRangeBase shehui_startCell = worksheet.Cells["K2"];
                ExcelRangeBase shehui_endCell = worksheet.Cells["L2"];
                int shehui_startRow = shehui_startCell.Start.Row;
                int shehui_startCol = shehui_startCell.Start.Column;
                int shehui_endRow = shehui_endCell.End.Row;
                int shehui_endCol = shehui_endCell.End.Column;
                //备注
                ExcelRangeBase beizhu_startCell = worksheet.Cells["O2"];
                ExcelRangeBase beizhu_endCell = worksheet.Cells["O3"];
                int beizhu_startRow = beizhu_startCell.Start.Row;
                int beizhu_startCol = beizhu_startCell.Start.Column;
                int beizhu_endRow = beizhu_endCell.End.Row;
                int beizhu_endCol = beizhu_endCell.End.Column;



                // 使用行列索引创建范围，并合并单元格  
                ExcelRangeBase Title = worksheet.Cells[Title_startRow, Title_startCol, Title_endRow, Title_endCol];
                ExcelRangeBase Num = worksheet.Cells[Num_startRow, Num_startCol, Num_endRow, Num_endCol];
                ExcelRangeBase Name = worksheet.Cells[Name_startRow, Name_startCol, Name_endRow, Name_endCol];
                ExcelRangeBase banji = worksheet.Cells[Class_startRow, Class_startCol, Class_endRow, Class_endCol];
                ExcelRangeBase StuID = worksheet.Cells[StuID_startRow, StuID_startCol, StuID_endRow, StuID_endCol];
                ExcelRangeBase Gread = worksheet.Cells[Gread_startRow, Gread_startCol, Gread_endRow, Gread_endCol];
                ExcelRangeBase GreadRank = worksheet.Cells[GreadRank_startRow, GreadRank_startCol, GreadRank_endRow, GreadRank_endCol];
                ExcelRangeBase sixiang = worksheet.Cells[sixiang_startRow, sixiang_startCol, sixiang_endRow, sixiang_endCol];
                ExcelRangeBase zhuanye = worksheet.Cells[zhuanye_startRow, zhuanye_startCol, zhuanye_endRow, zhuanye_endCol];
                ExcelRangeBase wenti = worksheet.Cells[wenti_startRow, wenti_startCol, wenti_endRow, wenti_endCol];
                ExcelRangeBase shehui = worksheet.Cells[shehui_startRow, shehui_startCol, shehui_endRow, shehui_endCol];
                ExcelRangeBase beizhu = worksheet.Cells[beizhu_startRow, beizhu_startCol, beizhu_endRow, beizhu_endCol];
                ExcelRangeBase hejifen1 = worksheet.Cells["E3"];
                ExcelRangeBase zhehefen1 = worksheet.Cells["F3"];
                ExcelRangeBase hejifen2 = worksheet.Cells["G3"];
                ExcelRangeBase zhehefen2 = worksheet.Cells["H3"];
                ExcelRangeBase hejifen3 = worksheet.Cells["I3"];
                ExcelRangeBase zhehefen3 = worksheet.Cells["J3"];
                ExcelRangeBase hejifen4 = worksheet.Cells["K3"];
                ExcelRangeBase zhehefen4 = worksheet.Cells["L3"];

                Title.Merge = Num.Merge = Name.Merge = StuID.Merge = banji.Merge = Gread.Merge = GreadRank.Merge = sixiang.Merge = zhuanye.Merge = wenti.Merge = shehui.Merge = beizhu.Merge = true;
                // 为合并后的单元格设置值（这将只显示在合并范围的左上角单元格中）  
                Title.Value = "XXXXX系XXX班综合素质测评统计表";
                Num.Value = "序号";
                Name.Value = "姓名";
                StuID.Value = "学号";
                banji.Value = "专业班级";
                Gread.Value = "总分";
                GreadRank.Value = "总分排名";
                sixiang.Value = "思想品德测评成绩";
                zhuanye.Value = "专业学习测评成绩";
                wenti.Value = "文体活动测评成绩";
                shehui.Value = "社会服务测评成绩";
                beizhu.Value = "备注";
                hejifen1.Value = "合计分";
                zhehefen1.Value = "折合分";
                hejifen2.Value = "合计分";
                zhehefen2.Value = "折合分";
                hejifen3.Value = "合计分";
                zhehefen3.Value = "折合分";
                hejifen4.Value = "合计分";
                zhehefen4.Value = "折合分";


                Title.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                Title.Style.Font.Size = 20; // 设置字号为20  
                Title.Style.Font.Bold = true;
                banji.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                banji.Style.Font.Size = 11; // 设置字号为20  
                banji.Style.Font.Bold = true;
                Name.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                Name.Style.Font.Size = 11; // 设置字号为11  
                Name.Style.Font.Bold = true;
                Num.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                Num.Style.Font.Size = 11; // 设置字号为11  
                Num.Style.Font.Bold = true;
                StuID.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                StuID.Style.Font.Size = 11; // 设置字号为11  
                StuID.Style.Font.Bold = true;
                Gread.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                Gread.Style.Font.Size = 11; // 设置字号为11  
                Gread.Style.Font.Bold = true;
                GreadRank.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                GreadRank.Style.Font.Size = 11; // 设置字号为11  
                GreadRank.Style.Font.Bold = true;
                sixiang.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                sixiang.Style.Font.Size = 11; // 设置字号为11  
                sixiang.Style.Font.Bold = true;
                zhuanye.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhuanye.Style.Font.Size = 11; // 设置字号为11  
                zhuanye.Style.Font.Bold = true;
                wenti.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                wenti.Style.Font.Size = 11; // 设置字号为11  
                wenti.Style.Font.Bold = true;
                shehui.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                shehui.Style.Font.Size = 11; //; 设置字号为11  
                shehui.Style.Font.Bold = true;
                beizhu.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                beizhu.Style.Font.Size = 11; // 设置字号为11  
                beizhu.Style.Font.Bold = true;
                hejifen1.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                hejifen1.Style.Font.Size = 11; // 设置字号为11  
                hejifen1.Style.Font.Bold = true;
                zhehefen1.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhehefen1.Style.Font.Size = 11; // 设置字号为11  
                zhehefen1.Style.Font.Bold = true;
                hejifen2.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                hejifen2.Style.Font.Size = 11; // 设置字号为11  
                hejifen2.Style.Font.Bold = true;
                zhehefen2.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhehefen2.Style.Font.Size = 11; // 设置字号为11  
                zhehefen2.Style.Font.Bold = true;
                hejifen3.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                hejifen3.Style.Font.Size = 11; // 设置字号为11  
                hejifen3.Style.Font.Bold = true;
                zhehefen3.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhehefen3.Style.Font.Size = 11; // 设置字号为11  
                zhehefen3.Style.Font.Bold = true;
                hejifen4.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                hejifen4.Style.Font.Size = 11; // 设置字号为11  
                hejifen4.Style.Font.Bold = true;
                zhehefen4.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhehefen4.Style.Font.Size = 11; // 设置字号为11  
                zhehefen4.Style.Font.Bold = true;

                Title.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Num.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Name.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                StuID.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                banji.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Gread.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                GreadRank.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sixiang.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhuanye.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wenti.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                shehui.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                beizhu.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hejifen1.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhehefen1.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hejifen2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhehefen2.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hejifen3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhehefen3.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                hejifen4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhehefen4.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                Title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Num.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Name.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                StuID.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                banji.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Gread.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                GreadRank.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sixiang.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhuanye.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wenti.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                shehui.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                beizhu.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hejifen1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhehefen1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hejifen2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhehefen2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hejifen3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhehefen3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                hejifen4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhehefen4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;


                // 保存更改到文件  
                package.Save();
            }
            /*
             * 
             * 
             * 创建加分表格
             * 
             * 
             * 
             */

            //让所有的空格子填上0
            using (var package = new ExcelPackage(FilePath))
            {
                // 获取第一个工作表  
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // 遍历所有行和列  
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        // 获取单元格  
                        ExcelRange cell = worksheet.Cells[row, col];

                        // 检查单元格是否为空（即没有值）  
                        if (cell.Value == null || cell.Value.ToString() == "")
                        {
                            // 如果为空，则设置为0  
                            cell.Value = 0;
                        }
                    }
                }

                // 保存更改  
                package.Save();
            }
            using (var package = new ExcelPackage(FilePath))
            {
                // 获取第一个工作表  
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                // 遍历所有行和列  
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        // 获取单元格  
                        ExcelRange cell = worksheet.Cells[row, col];

                        // 检查单元格是否为空（即没有值）  
                        if (cell.Value == null || cell.Value.ToString() == "")
                        {
                            // 如果为空，则设置为0  
                            cell.Value = 0;
                        }
                    }
                }

                // 保存更改  
                package.Save();
            }
            using (var package = new ExcelPackage(FilePath))
            {
                // 获取第一个工作表  
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];

                // 遍历所有行和列  
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        // 获取单元格  
                        ExcelRange cell = worksheet.Cells[row, col];

                        // 检查单元格是否为空（即没有值）  
                        if (cell.Value == null || cell.Value.ToString() == "")
                        {
                            // 如果为空，则设置为0  
                            cell.Value = 0;
                        }
                    }
                }

                // 保存更改  
                package.Save();
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))    //将原始工作表的成绩读取
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // 获取第一个工作表  

                List<KeyValuePair<string, int>> sumMappingsD = new List<KeyValuePair<string, int>>();
                List<KeyValuePair<string, int>> sumMappingsF = new List<KeyValuePair<string, int>>();
                List<KeyValuePair<string, int>> sumMappingsH = new List<KeyValuePair<string, int>>();
                List<KeyValuePair<string, int>> sumMappingsJ = new List<KeyValuePair<string, int>>();

                // 遍历B列和D列的所有单元格  
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)   //row = 2：从第二行开始计算，第一行为标题
                {
                    string cellDValue = worksheet.Cells[row, 4].Value?.ToString();
                    string cellFValue = worksheet.Cells[row, 6].Value?.ToString();
                    string cellHValue = worksheet.Cells[row, 8].Value?.ToString();
                    string cellJValue = worksheet.Cells[row, 10].Value?.ToString();
                    string cellBValue_Name = worksheet.Cells[row, 2].Value?.ToString();

                    if (!string.IsNullOrEmpty(cellBValue_Name))      //从第二行开始，如果读取到了空行，结束循环
                    {
                        // 使用正则表达式匹配字符串中的数字，并求和  
                        MatchCollection matchesD = Regex.Matches(cellDValue, @"\d+");
                        MatchCollection matchesF = Regex.Matches(cellFValue, @"\d+");
                        MatchCollection matchesH = Regex.Matches(cellHValue, @"\d+");
                        MatchCollection matchesJ = Regex.Matches(cellJValue, @"\d+");

                        int sumD = matchesD.Cast<Match>().Select(m => int.Parse(m.Value)).Sum();
                        int sumF = matchesF.Cast<Match>().Select(m => int.Parse(m.Value)).Sum();
                        int sumH = matchesH.Cast<Match>().Select(m => int.Parse(m.Value)).Sum();
                        int sumJ = matchesJ.Cast<Match>().Select(m => int.Parse(m.Value)).Sum();

                        // 将D列的数据与求和结果构成映射关系，并存入列表中  
                        sumMappingsD.Add(new KeyValuePair<string, int>(cellDValue, sumD));
                        System.Diagnostics.Debug.WriteLine($"当前运行到第 {row} 行了");
                        worksheet.Cells[row, 4].Value = sumD;
                        System.Diagnostics.Debug.WriteLine($"sumD = {sumD} ");
                        // 将F列的数据与求和结果构成映射关系，并存入列表中  
                        sumMappingsF.Add(new KeyValuePair<string, int>(cellFValue, sumF));
                        worksheet.Cells[row, 6].Value = sumF;
                        System.Diagnostics.Debug.WriteLine($"sumF = {sumF} ");
                        // 将H列的数据与求和结果构成映射关系，并存入列表中  
                        sumMappingsH.Add(new KeyValuePair<string, int>(cellHValue, sumH));
                        worksheet.Cells[row, 8].Value = sumH;
                        System.Diagnostics.Debug.WriteLine($"sumH = {sumH} ");
                        // 将J列的数据与求和结果构成映射关系，并存入列表中  
                        sumMappingsJ.Add(new KeyValuePair<string, int>(cellJValue, sumJ));
                        worksheet.Cells[row, 10].Value = sumJ;
                        System.Diagnostics.Debug.WriteLine($" sumJ = {sumJ} ");
                        System.Diagnostics.Debug.WriteLine($"第 {row} 结束");
                    }
                }
                int rowD, rowF, rowH, rowJ, row_now;
                rowD = rowF = rowH = rowJ = 0;
                row_now = 2;
                // 打印或处理映射关系列表  
                foreach (var mappingD in sumMappingsD)
                {
                    System.Diagnostics.Debug.WriteLine($"D列值: {mappingD.Key}, 数字之和: {mappingD.Value}");     //思想品德
                }
                foreach (var mappingF in sumMappingsF)
                {
                    System.Diagnostics.Debug.WriteLine($"F列值: {mappingF.Key}, 数字之和: {mappingF.Value}");     //专业学习
                }
                foreach (var mappingH in sumMappingsH)
                {
                    System.Diagnostics.Debug.WriteLine($"H列值: {mappingH.Key}, 数字之和: {mappingH.Value}");     //文体活动
                }
                foreach (var mappingJ in sumMappingsJ)
                {
                    System.Diagnostics.Debug.WriteLine($"J列值: {mappingJ.Key}, 数字之和: {mappingJ.Value}");     //社会服务
                }
                System.Diagnostics.Debug.WriteLine($"rowD={rowD},rowF={rowF},rowH={rowH},rowJ={rowJ},");
                package.SaveAs(FilePath);
            }


            //添加功能：班上人数是否正确，如果不正确就重新检查表格啥的

            //文本切换  
            FileInfo excelFile = new FileInfo(FilePath);
            // 加载Excel包  
            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                // 获取第二个工作表  
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                // 要搜索的列索引（D、F、H、J）  
                int[] columnIndices = { 4, 6, 8, 10 };

                // 遍历要搜索的列  
                foreach (var colIndex in columnIndices)
                {
                    // 遍历该列的每个单元格  
                    for (int rowIndex = 1; rowIndex <= worksheet.Dimension.End.Row; rowIndex++)
                    {
                        ExcelRangeBase cell = worksheet.Cells[rowIndex, colIndex];
                        if (cell.Value != null)
                        {
                            // 获取单元格内容，并分割成字符串数组  
                            string[] stringsInCell = cell.Value.ToString().Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                            // 用于存储替换后的字符串  
                            List<string> updatedStrings = new List<string>();

                            // 遍历单元格内的每个字符串  
                            foreach (var str in stringsInCell)
                            {
                                // 搜索column1Data中的匹配项  
                                for (int i = 0; i < column1Data.Count; i++)
                                {
                                    if (str.Contains(column1Data[i]))
                                    {
                                        updatedStrings.Add(column1Data[i] + "+" + column3Data[i]);
                                        break; // 找到匹配项后跳出循环  
                                    }
                                }

                                //// 如果没有找到匹配项，则保留原字符串  
                                //if (!updatedStrings.Contains(str))
                                //{
                                //    updatedStrings.Add(str);
                                //}
                            }

                            // 将更新后的字符串数组连接回一个字符串，并用逗号加空格分隔  
                            string updatedCellValue = string.Join(", ", updatedStrings);

                            // 将更新后的值写回单元格  
                            cell.Value = updatedCellValue;
                        }
                    }
                }

                // 保存更改后的Excel文件
                package.SaveAs(FilePath);
            }
            System.Diagnostics.Debug.WriteLine("完成替换");


            //总计算
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  //原始表
                ExcelWorksheet worksheet_zongfen = package.Workbook.Worksheets[3];  //总分表

                int lastRowWithData_yuanshi = worksheet.Dimension.End.Row; // 获取总分工作表的最大行数  
                System.Diagnostics.Debug.WriteLine($"原始表的最大一行是：{lastRowWithData_yuanshi}");

                //先把合计分数全都移到总分表中
                // 定义源范围：从D列第二行到最后一个有数据的行  
                ExcelRangeBase sourceRangeD = worksheet.Cells["D2:D" + lastRowWithData_yuanshi];    //思想品德
                ExcelRangeBase sourceRangeF = worksheet.Cells["F2:F" + lastRowWithData_yuanshi];    //专业学习
                ExcelRangeBase sourceRangeH = worksheet.Cells["H2:H" + lastRowWithData_yuanshi];    
                ExcelRangeBase sourceRangeJ = worksheet.Cells["J2:J" + lastRowWithData_yuanshi];
                ExcelRangeBase sourceRange_name = worksheet.Cells["B2:B" + lastRowWithData_yuanshi];  //姓名
                ExcelRangeBase sourceRange_ID = worksheet.Cells["C2:C" + lastRowWithData_yuanshi];  //学号
                // 定义目标范围的起始单元格（例如Sheet2的A1）  
                ExcelRangeBase targetRangeD = worksheet_zongfen.Cells["E4"];//思想品德合计分
                ExcelRangeBase targetRangeF = worksheet_zongfen.Cells["G4"];//专业学习合计分
                ExcelRangeBase targetRangeH = worksheet_zongfen.Cells["I4"];//
                ExcelRangeBase targetRangeJ = worksheet_zongfen.Cells["K4"];//社会服务合计分
                ExcelRangeBase targetRange_name = worksheet_zongfen.Cells["B4"];//姓名栏
                ExcelRangeBase targetRange_ID = worksheet_zongfen.Cells["C4"];//学号
                // 复制源范围到目标范围  
                sourceRangeD.Copy(targetRangeD);
                sourceRangeF.Copy(targetRangeF);
                sourceRangeH.Copy(targetRangeH);
                sourceRangeJ.Copy(targetRangeJ);
                sourceRange_name.Copy(targetRange_name);
                sourceRange_ID.Copy(targetRange_ID);

                package.SaveAs(FilePath);

                //在总分表中处理合计分到折合分
                int row_now = 4;  //行
                int col_now_chengji = 4;

                int col_xuefen, row_xuekechengji, col_xuekechengji,row_guake;
                row_guake = 3;
                row_xuekechengji = 3;
                col_xuekechengji = 4;
                col_xuefen = 4;  //从第四列开始向右遍历

                ExcelWorksheet worksheet1 = package.Workbook.Worksheets[3];  //总分表
                ExcelWorksheet worksheet_chengji = package.Workbook.Worksheets[2];  //成绩表

                int lastRowWithData_zongfen = worksheet1.Dimension.End.Row; // 获取总分工作表的最大行数 
                System.Diagnostics.Debug.WriteLine($"总分表的最大一行是：{lastRowWithData_zongfen}");

                int maxColWithContent_chengji = worksheet_chengji.Dimension.End.Column; // 获取总分工作表的最大行数 
                System.Diagnostics.Debug.WriteLine($"成绩的最大一列是：{maxColWithContent_chengji}");

                int maxColWithContent = 15;
                System.Diagnostics.Debug.WriteLine($"总分表的最大一列是：{maxColWithContent}");

                int lastRowWithData_chengji = worksheet_chengji.Dimension.End.Row; // 获取成绩工作表的最大行数  
                System.Diagnostics.Debug.WriteLine($"成绩的最大一行是：{lastRowWithData_chengji}");


                worksheet1.Column(2).Width = 10;//设置列宽
                worksheet1.Column(3).Width = 12;//设置列宽

                

                int col;
                col = 5;
                while (col != 13) //先处理列，5-7-9-11，如果为13，那么跳出循环
                {
                    while (row_now <= lastRowWithData_zongfen)
                    {
                        // 获取当前列的单元格  
                        ExcelRangeBase cell = worksheet1.Cells[row_now, col];  //初始合计：第4行第5列
                        if (col == 5)//处理文体活动与社会服务：占比均为10%
                        {
                            // 检查单元格是否为空,如果为空，那么换行检测
                            if (cell.Value != null && double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double cellValue))
                            {
                                System.Diagnostics.Debug.WriteLine("总分表：当前在第" + row_now + "行，思想品德列");

                                // 读取单元格的值（假设它是数字）并进行计算  
                                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double value))
                                {
                                    // 在这里执行你的计算，例如将值乘以2  
                                    double calculatedValue = (value + 80) * 0.15;

                                    // 将计算结果写回原来的单元格  
                                    if(value + 80 <= 100)
                                    {
                                        Math.Round(value, 2);
                                        worksheet1.Cells[row_now, col].Value = value + 80;
                                        worksheet1.Cells[row_now, col + 1].Value = calculatedValue; 
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format =  "#,##0.00";
                                        //worksheet1.Cells[].Style.Font.Size = worksheet1.Cells[].Style.Font.Size = 11; // 设置字号为20  
                                        //worksheet1.Cells[].Style.HorizontalAlignment = worksheet1.Cells[].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        //worksheet1.Cells[].Style.VerticalAlignment = worksheet1.Cells[].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        //worksheet1.Cells[].Style.Numberformat.Format = worksheet1.Cells[].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    else
                                    {
                                        worksheet1.Cells[row_now, col].Value = 100;
                                        worksheet1.Cells[row_now, col+ 1].Value = 15;
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";

                                    }
                                }
                            }
                        }
                        row_now++;
                        package.SaveAs(FilePath);
                    }
                    row_now = 4;    //回到第一行
                    col = col + 2;   //前往下一个类别（跨过折合分）
                    while (row_now <= lastRowWithData_zongfen)
                    {
                        ExcelRangeBase cell = worksheet1.Cells[row_now, col];  //初始合计：第4行第7列
                        if (col == 7)//处理专业学习：占比65%
                        {
                            //从两个表中获取所需参数
                            //ExcelRangeBase cell_xuefen = worksheet_chengji.Cells[2, col_xuekechengji];  //获取当前学科成绩的学分
                            ExcelRangeBase cell_last = worksheet_chengji.Cells[row_now - 1, maxColWithContent_chengji];  //从“成绩”获取当前最后一步所需的，已经除以总学分的成绩
                            ExcelRangeBase cell_fujia = worksheet1.Cells[row_now, col];  //获取“总表”已经拥有的附加分数

                            //开始做减法
                            //1.double化两个所需的参数
                            double.TryParse(cell_last.Value.ToString(), out double value_last);      //??????
                            double.TryParse(cell_fujia.Value.ToString(), out double value_fujia);
                            //2.相减
                            double calculatedValue = value_last - value_fujia;
                            worksheet1.Cells[row_now, col].Value = calculatedValue;
                            worksheet1.Cells[row_now, col + 1].Value = calculatedValue * 0.65;
                            worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                            worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";
                        }
                        row_now++;
                        package.SaveAs(FilePath);
                    }
                    row_now = 4;    //回到第一行
                    col = col + 2;   //前往下一个类别（跨过折合分）

                    while (row_now <= lastRowWithData_zongfen)
                    {
                        ExcelRangeBase cell = worksheet1.Cells[row_now, col];  //初始合计：第4行第9列
                        if (col == 9)//处理文体活动：占比均为10%
                        {
                            // 检查单元格是否为空,如果为空，那么换行检测
                            if (cell.Value != null && double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double cellValue))
                            {
                                System.Diagnostics.Debug.WriteLine("总分表：当前在第" + row_now + "行，文体活动列");

                                // 读取单元格的值（假设它是数字）并进行计算  
                                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double value))
                                {
                                    // 在这里执行你的计算，例如将值乘以2  
                                    double calculatedValue = (value + 80) * 0.1;

                                    // 将计算结果写回原来的单元格  
                                    if (value + 80 <= 100)
                                    {
                                        worksheet1.Cells[row_now, col].Value = value + 80;
                                        worksheet1.Cells[row_now, col + 1].Value = calculatedValue;
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";

                                    }
                                    else
                                    {
                                        worksheet1.Cells[row_now, col].Value = 100;
                                        worksheet1.Cells[row_now, col + 1].Value = 10;
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";
                                    }
                                }
                            }
                        }
                        row_now++;
                    }




                    row_now = 4;    //回到第一行
                    col = col + 2;   //前往下一个类别（跨过折合分）
                    while (row_now <= lastRowWithData_zongfen)
                    {
                        ExcelRangeBase cell = worksheet1.Cells[row_now, col];  //初始合计：第4行第11列
                        if (col == 11)//社会服务：占比均为10%
                        {
                            // 检查单元格是否为空,如果为空，那么换行检测
                            if (cell.Value != null && double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double cellValue))
                            {
                                System.Diagnostics.Debug.WriteLine("总分表：当前在第" + row_now + "行，社会服务列");

                                // 读取单元格的值（假设它是数字）并进行计算  
                                if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double value))
                                {
                                    // 在这里执行你的计算，例如将值乘以2  
                                    double calculatedValue = (value + 80) * 0.1;

                                    // 将计算结果写回原来的单元格  
                                    if (value + 80 <= 100)
                                    {
                                        worksheet1.Cells[row_now, col].Value = value + 80;
                                        worksheet1.Cells[row_now, col + 1].Value = calculatedValue;
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";
                                    }
                                    else
                                    {
                                        worksheet1.Cells[row_now, col].Value = 100;
                                        worksheet1.Cells[row_now, col+1].Value = 10;
                                        worksheet1.Cells[row_now, col].Style.Font.Size = worksheet1.Cells[row_now, col + 1].Style.Font.Size = 11; // 设置字号为20  
                                        worksheet1.Cells[row_now, col].Style.HorizontalAlignment = worksheet1.Cells[row_now, col + 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.VerticalAlignment = worksheet1.Cells[row_now, col + 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                        worksheet1.Cells[row_now, col].Style.Numberformat.Format = worksheet1.Cells[row_now, col + 1].Style.Numberformat.Format = "#,##0.00";
                                    }
                                }
                            }
                        }
                        row_now++;
                    }
                    row_now = 4;    //回到第一行
                    col = col + 2;   //当前col == 13
                }
                package.SaveAs(FilePath);
                //分数计算完成，获取所以类别的折合分
                int row_zongfen, col_zongfen, sum_zongfen;
                row_zongfen = 4;
                col_zongfen = 6;
                sum_zongfen = 0;
                while (row_zongfen <= lastRowWithData_zongfen)
                {
                    ExcelRange cell_sixiangzhehe = worksheet1.Cells[row_zongfen, 6];    //获取思想折合分
                    System.Diagnostics.Debug.WriteLine(row_zongfen);
                    ExcelRange cell_zhuanyezhehe = worksheet1.Cells[row_zongfen, 8];    //获取专业折合分
                    System.Diagnostics.Debug.WriteLine(row_zongfen);
                    ExcelRange cell_wentizhehe2 = worksheet1.Cells[row_zongfen, 10];    //获取文体折合分
                    System.Diagnostics.Debug.WriteLine(row_zongfen);
                    ExcelRange cell_shehuizhehe = worksheet1.Cells[row_zongfen, 12];    //获取社会折合分
                    System.Diagnostics.Debug.WriteLine(row_zongfen);
                    double.TryParse(cell_sixiangzhehe.Value.ToString(), out double value_sixiangzhehe);
                    double.TryParse(cell_zhuanyezhehe.Value.ToString(), out double value_zhuanyezhehe);
                    double.TryParse(cell_wentizhehe2.Value.ToString(), out double value_wentizhehe);  //??????
                    double.TryParse(cell_shehuizhehe.Value.ToString(), out double value_shehuizhehe);
                    worksheet1.Cells[row_zongfen, 13].Value = value_sixiangzhehe + value_zhuanyezhehe + value_wentizhehe + value_shehuizhehe;
                    worksheet1.Cells[row_zongfen, 13].Style.Font.Size = 11; // 设置字号为20  
                    worksheet1.Cells[row_zongfen, 13].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet1.Cells[row_zongfen, 13].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet1.Cells[row_zongfen, 13].Style.Numberformat.Format = "#,##0.00";
                    worksheet1.Cells[row_zongfen, 13].Style.ShrinkToFit = true;
                    row_zongfen++;

                }

                package.SaveAs(FilePath);

                //开始处理备注文本
                //将文本生成里的相关单元格内容存入列表中
            }
            using (ExcelPackage Package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet Worksheet_wenben = Package.Workbook.Worksheets[1];   //打开文本生成工作表
                ExcelWorksheet Worksheet_zongfen = Package.Workbook.Worksheets[3];   //打开总分工作表

                int maxRow_wenben = Worksheet_wenben.Dimension.End.Row; // 获取文本生成工作表的最大行数  
                System.Diagnostics.Debug.WriteLine($"文本生成的最大一行是：{maxRow_wenben}");
                int maxCol_wenben = Worksheet_wenben.Dimension.End.Column;  // 获取文本生成工作表的最大列数  
                System.Diagnostics.Debug.WriteLine($"文本生成的最大一列是：{maxCol_wenben}");

                int maxRow_zongfen = Worksheet_zongfen.Dimension.End.Row; // 获取总分工作表的最大行数  
                System.Diagnostics.Debug.WriteLine($"总分的最大一行是：{maxRow_zongfen}");
                int maxCol_zongfen = Worksheet_zongfen.Dimension.End.Column; // 获取总分工作表的最大列数  
                System.Diagnostics.Debug.WriteLine($"总分的最大一列是：{maxCol_zongfen}");

                int row_zongfen, col_zongfen, row_wenben, col_wenben;
                row_zongfen = 4;
                col_zongfen = 0;
                row_wenben = 2;
                col_wenben = 4;
                while(row_wenben <= maxRow_wenben)
                {
                    ExcelRange cell_wenben_sixiang = Worksheet_wenben.Cells[row_wenben, 4];    //获取思想品德备注
                    ExcelRange cell_wenben_zhuanye = Worksheet_wenben.Cells[row_wenben, 6];    //获取专业学习备注
                    ExcelRange cell_wenben_wenti = Worksheet_wenben.Cells[row_wenben, 8];    //获文体活动想备注
                    ExcelRange cell_wenben_shehui = Worksheet_wenben.Cells[row_wenben, 10];    //获取社会服务备注


                    string sixiang = cell_wenben_sixiang.Value.ToString();
                    string zhuanye = cell_wenben_zhuanye.Value.ToString();
                    string wenti = cell_wenben_wenti.Value.ToString();     //??????
                    string shehui = cell_wenben_shehui.Value.ToString();
                    if (Worksheet_wenben.Cells[row_wenben, 4].Value != null)
                    {
                        Worksheet_zongfen.Cells[row_zongfen, 15].Value = "思想品德：" + sixiang;
                        Console.WriteLine($"第{row_wenben}行思想品德");
                    }
                    if (Worksheet_wenben.Cells[row_wenben, 6].Value != null)
                    {
                        Worksheet_zongfen.Cells[row_zongfen, 15].Value = Worksheet_zongfen.Cells[row_zongfen, 15].Value + "专业学习：" + zhuanye;
                        Console.WriteLine($"第{row_wenben}行专业学习");
                    }
                    if (Worksheet_wenben.Cells[row_wenben, 8].Value != null)
                    {
                        Worksheet_zongfen.Cells[row_zongfen, 15].Value = Worksheet_zongfen.Cells[row_zongfen, 15].Value + "文体活动：" + wenti;
                        Console.WriteLine($"第{row_wenben}行文体活动");
                    }
                    if (Worksheet_wenben.Cells[row_wenben, 10].Value != null)
                    {
                        Worksheet_zongfen.Cells[row_zongfen, 15].Value = Worksheet_zongfen.Cells[row_zongfen, 15].Value + "社会服务：" + shehui;
                        Console.WriteLine($"第{row_wenben}行社会服务");
                    }
                    row_zongfen++;
                    row_wenben++;
                    Package.SaveAs(FilePath);
                }
                Package.SaveAs(FilePath);
            }
            using (var package = new ExcelPackage(new FileInfo(FilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[3]; // 假设我们要操作的是第一个工作表  

                // 读取所有行的数据，并创建一个包含所有行的列表  
                var rows = new List<Dictionary<int, object>>();
                int startRow = 3; // 假设第3行是标题行，从第二行开始读取数据  
                int endRow = worksheet.Dimension.End.Row;

                for (int row = startRow; row <= endRow; row++)
                {
                    var rowData = new Dictionary<int, object>();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        rowData[col] = worksheet.Cells[row, col].Value;
                    }
                    rows.Add(rowData);
                }

                // 根据N列的数字大小进行排序  
                rows = rows
                   .OrderByDescending(r =>
                   {
                       if (r.ContainsKey(13) && r[13] != null) // 确保M列的键存在且值不为null  
                       {
                           // 尝试转换，如果转换失败则返回double.MaxValue（或其他适当的值）  
                           if (double.TryParse(r[13].ToString(), out double result))
                           {
                               return result; // 假设M列是数字，转换为double进行排序  
                           }
                           else
                           {
                               return double.MaxValue; // 如果无法转换，则放在最后  
                           }
                       }
                       return double.MaxValue;  // 如果M列的键不存在或值为null，则放在最后  
                   })
                    .ToList();

                // 清空除标题行外的所有行数据  
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        worksheet.Cells[row, col].Value = null;
                    }
                }

                // 将排序后的数据写回Excel中  
                int rowToWrite = startRow - 1; // 标题行不需要重新写，所以从第一行数据开始写  
                foreach (var rowData in rows)
                {
                    rowToWrite++;
                    foreach (var kvp in rowData)
                    {
                        worksheet.Cells[rowToWrite, kvp.Key].Value = kvp.Value;
                    }
                }




                int xuhao = 0;
                //建立序号
                for (int i = 4; i <= rowToWrite; i++)
                {
                    xuhao++;
                    worksheet.Cells[i, 1].Value = xuhao;
                    worksheet.Cells[i, 1].Style.Font.Name = "等线"; // 设置字体为“微软雅黑”   
                    worksheet.Cells[i, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[i, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Column(1).Width = 10;//设置列宽
                    worksheet.Cells[i, 14].Value = xuhao;
                    worksheet.Cells[i, 14].Style.Font.Name = "等线"; // 设置字体为“微软雅黑”   
                    worksheet.Cells[i, 14].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[i, 14].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Column(1).Width = 10;//设置列宽
                    //Title.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    //Title.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                }
                ExcelRangeBase sixiang_startCell = worksheet.Cells["E2"];
                ExcelRangeBase sixiang_endCell = worksheet.Cells["F2"];
                int sixiang_startRow = sixiang_startCell.Start.Row;
                int sixiang_startCol = sixiang_startCell.Start.Column;
                int sixiang_endRow = sixiang_endCell.End.Row;
                int sixiang_endCol = sixiang_endCell.End.Column;
                //专业学习
                ExcelRangeBase zhuanye_startCell = worksheet.Cells["G2"];
                ExcelRangeBase zhuanye_endCell = worksheet.Cells["H2"];
                int zhuanye_startRow = zhuanye_startCell.Start.Row;
                int zhuanye_startCol = zhuanye_startCell.Start.Column;
                int zhuanye_endRow = zhuanye_endCell.End.Row;
                int zhuanye_endCol = zhuanye_endCell.End.Column;
                //文体活动
                ExcelRangeBase wenti_startCell = worksheet.Cells["I2"];
                ExcelRangeBase wenti_endCell = worksheet.Cells["J2"];
                int wenti_startRow = wenti_startCell.Start.Row;
                int wenti_startCol = wenti_startCell.Start.Column;
                int wenti_endRow = wenti_endCell.End.Row;
                int wenti_endCol = wenti_endCell.End.Column;
                //社会服务
                ExcelRangeBase shehui_startCell = worksheet.Cells["K2"];
                ExcelRangeBase shehui_endCell = worksheet.Cells["L2"];
                int shehui_startRow = shehui_startCell.Start.Row;
                int shehui_startCol = shehui_startCell.Start.Column;
                int shehui_endRow = shehui_endCell.End.Row;
                int shehui_endCol = shehui_endCell.End.Column;
                ExcelRangeBase sixiang = worksheet.Cells[sixiang_startRow, sixiang_startCol, sixiang_endRow, sixiang_endCol];
                ExcelRangeBase zhuanye = worksheet.Cells[zhuanye_startRow, zhuanye_startCol, zhuanye_endRow, zhuanye_endCol];
                ExcelRangeBase wenti = worksheet.Cells[wenti_startRow, wenti_startCol, wenti_endRow, wenti_endCol];
                ExcelRangeBase shehui = worksheet.Cells[shehui_startRow, shehui_startCol, shehui_endRow, shehui_endCol];
                sixiang.Merge = zhuanye.Merge = wenti.Merge = shehui.Merge = true;
                sixiang.Value = "思想品德测评成绩";
                zhuanye.Value = "专业学习测评成绩";
                wenti.Value = "文体活动测评成绩";
                shehui.Value = "社会服务测评成绩";
                sixiang.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                sixiang.Style.Font.Size = 11; // 设置字号为11  
                sixiang.Style.Font.Bold = true;
                zhuanye.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                zhuanye.Style.Font.Size = 11; // 设置字号为11  
                zhuanye.Style.Font.Bold = true;
                wenti.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                wenti.Style.Font.Size = 11; // 设置字号为11  
                wenti.Style.Font.Bold = true;
                shehui.Style.Font.Name = "微软雅黑"; // 设置字体为“微软雅黑”  
                shehui.Style.Font.Size = 11; //; 设置字号为11  
                shehui.Style.Font.Bold = true;
                sixiang.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                zhuanye.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                wenti.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                shehui.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sixiang.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                zhuanye.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                wenti.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                shehui.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                worksheet.Cells[4, 4].Value = "XX专业XX班";
                worksheet.Cells[4, 4].Style.Font.Name = "等线"; // 设置字体为“微软雅黑”  
                worksheet.Cells[4, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[4, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Column(4).Width = 13;//设置列宽

                MessageBox.Show("完成！");
                package.SaveAs(new FileInfo(FilePath)); // 如果需要保存到新的文件，可以传递新的文件路径给SaveAs方法  
            }
        }
    }
}