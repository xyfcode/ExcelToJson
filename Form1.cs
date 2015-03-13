using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace 读取Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnInput_Click(object sender, EventArgs e)
        {
            string fileName = textBox1.Text;           

            if (fileName.EndsWith("xlsx") || fileName.EndsWith("xls"))
            {
               

                if (File.Exists(fileName))
                {

                    //根据路径通过已存在的excel来创建HSSFWorkbook，即整个excel文档
                    FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                    XSSFWorkbook workbook = new XSSFWorkbook(file);


                    WriteJson(workbook);
                    MessageBox.Show("导出成功");
                }
            }
            else 
            {
                textBox1.Text = "";
                MessageBox.Show("请输入Excel格式的文件");
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr=openFileDialog1.ShowDialog();
            if (dr.ToString() == "OK")
            {                
                textBox1.Text = openFileDialog1.FileName;                               
            }
            else 
            {
                MessageBox.Show("输入的地址不存在");
            }      
        }

        private void WriteJson(XSSFWorkbook workbook)
        {
            //服务端表头位置
            var server_title = 1;
            //服务端数据开始行数
            var server_num = 4;

            //获取excel的第一个sheet

            string sheet_name=workbook.GetSheetName(0);
             

            ISheet sheet = workbook.GetSheetAt(0);

            try
            {

                string txtPath = @".\json\" + sheet_name.ToLower() + ".json";
                FileStream aFile = new FileStream(txtPath, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(aFile);
                sw.Write("{\r\n");
                sw.Write("   \"" + sheet_name.Trim().ToUpper() + "\":{\r\n");


                //获取sheet的第二行，服务端用的表头
                IRow titleRow = sheet.GetRow(server_title);

                //一行最后一个方格的编号 即总的列数
                int cellCount = titleRow.LastCellNum;

                //for (int i = titleRow.FirstCellNum; i < cellCount; i++)
                //{
                    
                //    //titleRow.GetCell(i).StringCellValue;                    
                //}

                //最后一行
                int rowCount = sheet.LastRowNum;
                

                //遍历行
                for (int i = server_num; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        MessageBox.Show("可能有异常，请对比数据！");
                        break;
                    }

                    string str_write = "        \"" + (i - server_num + 1) + "\":{\r\n";
                    sw.Write(str_write);


                    string str_write2 = "";
                    //遍历该行的列
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {

                        if (titleRow.GetCell(j)!=null && titleRow.GetCell(j).ToString().Length != 0)
                        {
                            string value = "";
                            
                            if (row.GetCell(j) != null)
                            {
                                value = row.GetCell(j).ToString().Trim();
                            }

                            string title = titleRow.GetCell(j).ToString().Trim();                          
                           
                            str_write2 += "          \"" + title + "\":\"" + value + "\",\r\n"; 

                        }                      
                           
                    }
                    
                    int end_str2 = str_write2.LastIndexOf(",");
                    if (end_str2 != -1) 
                    {
                        str_write2 = str_write2.Remove(end_str2, 1);
                        sw.Write(str_write2);
                    }
                    
                 

                   
                    if (i == rowCount)
                    {
                        sw.Write("        }\r\n");
                    }
                    else
                    {
                        sw.Write("        },\r\n");
                    }
                   
                    
                }




                sw.Write("   }\r\n");
                sw.Write("}\r\n");
                sw.Close();
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);                
                return;
            }

        }

        private void buttonGm_Click(object sender, EventArgs e)
        {
            string fileName = textBox1.Text;

            if (fileName.EndsWith("xlsx") || fileName.EndsWith("xls"))
            {


                if (File.Exists(fileName))
                {

                    //根据路径通过已存在的excel来创建HSSFWorkbook，即整个excel文档
                    FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read);

                    XSSFWorkbook workbook = new XSSFWorkbook(file);


                    WriteArrGM(workbook);
                    MessageBox.Show("导出成功");
                }
            }
            else
            {
                textBox1.Text = "";
                MessageBox.Show("请输入Excel格式的文件");
            }
        }


        private void WriteArrGM(XSSFWorkbook workbook)
        {
            //服务端表头位置
            var server_title = 1;
            //服务端数据开始行数
            var server_num = 4;

            //获取excel的第一个sheet

            string sheet_name = workbook.GetSheetName(0);


            ISheet sheet = workbook.GetSheetAt(0);

            try
            {

                string txtPath = @".\json\" + sheet_name.ToLower() + ".json";
                FileStream aFile = new FileStream(txtPath, FileMode.OpenOrCreate);
                StreamWriter sw = new StreamWriter(aFile);
                sw.Write("[\r\n");          


                //获取sheet的第二行，服务端用的表头
                IRow titleRow = sheet.GetRow(server_title);

                //一行最后一个方格的编号 即总的列数
                int cellCount = titleRow.LastCellNum;
             

                //最后一行
                int rowCount = sheet.LastRowNum;


                //遍历行
                for (int i = server_num; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);

                    if (row == null)
                    {
                        MessageBox.Show("可能有异常，请对比数据！");
                        break;
                    }


                    string str_write2 = "";
                    //遍历该行的列
                    for (int j = row.FirstCellNum; j < 2; j++)
                    {

                        if (titleRow.GetCell(j) != null)
                        {
                            string value = "";

                            if (row.GetCell(j) != null)
                            {
                                value = row.GetCell(j).ToString().Trim();
                            }

                            string title = titleRow.GetCell(j).ToString().Trim();

                            if (j == 0)
                            {
                                str_write2 += "{ id: '" + value + "',";
                            }
                            else
                            {
                                str_write2 += " text: '" + value + "', pid :0 },";
                            }

                            

                        }

                    }

                    int end_str2 = str_write2.LastIndexOf(",");
                    if (end_str2 != -1)
                    {
                        str_write2 = str_write2.Remove(end_str2, 1);
                        sw.Write(str_write2);
                    }




                    if (i == rowCount)
                    {
                        sw.Write("\r\n");
                    }
                    else
                    {
                        sw.Write(",\r\n");
                    }


                }




               
                sw.Write("]\r\n");
                sw.Close();
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }
        

       
    }
}
