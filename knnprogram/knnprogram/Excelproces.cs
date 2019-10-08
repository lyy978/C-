using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Configuration;
using System.Web;
using Microsoft.Office.Core;

namespace knnprogram
{
    public class ExcelEdit
    {
        public string mFilename;
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;

        public void Create()//创建一个Microsoft.Office.Interop.Excel对象
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(true);
        }
        public void Open(string FileName)//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            //wb = wbs.Open(FileName, 0, true, 5,"", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "t", false, false, 0, true,Type.Missing,Type.Missing);
            //wb = wbs.Open(FileName,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Microsoft.Office.Interop.Excel.XlPlatform.xlWindows,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing,Type.Missing);
            mFilename = FileName;
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[SheetName];
            return s;
        }
        public Microsoft.Office.Interop.Excel.Worksheet AddSheet(string SheetName)
        //添加一个工作表
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            s.Name = SheetName;
            return s;
        }
        public void Close()
        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
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
        public bool Save()
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
        public void SetCellValue(Microsoft.Office.Interop.Excel.Worksheet ws, int x, int y, object value)
        //ws：要设值的工作表     X行Y列     value   值
        {
            ws.Cells[x, y] = value;

        }
        public void SetCellValue(string ws, int x, int y, object value)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {

            GetSheet(ws).Cells[x, y] = value;
        }

        public List<string> ColumnDB = new List<string>();
        public List<int> trainX1 = new List<int>();
        public List<int> trainX2 = new List<int>();
        public List<int> trainX3 = new List<int>();
        public List<double> trainX4 = new List<double>();
        public List<double> trainX5 = new List<double>();
        public List<double> trainY1 = new List<double>();
        public List<double> trainY2 = new List<double>();
        public List<double> trainY3 = new List<double>();


        public void getColumnInt(ExcelEdit ed)
        {

            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //选择指定工作表
                                                                                 //获取选选择的工作表
                                                                                 // Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);//方法一：指定工作表名称读取
                                                                                 //Worksheet ws = (Worksheet)openwb.Worksheets.get_Item(1);//方法二：通过工作表下标读取                                                                              
            int rows = worksheet.UsedRange.Rows.Count;                           //获取工作表中的行数  
            int columns = worksheet.UsedRange.Columns.Count;                     //获取工作表中的列数
            Console.WriteLine("请输入你要获取哪列数据");
            int column = Convert.ToInt16(Console.ReadLine());
            int m = 0;
            //提取对应行列的数据并将其存入数组中，按列读取
            for (int i = 2; i <= rows; i++)
            {
                int temp;
                string a = (worksheet.Cells[i, column]).Text.ToString();
                temp = Convert.ToInt32(a);
                //   Console.WriteLine("读取的数据:{0}" ,temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX1.Add(temp);
                m++;
            }

            //遍历数组
            foreach (string db in ColumnDB)
            {
                Console.WriteLine("列表中的数据" + db);//查看数组中的数据，测试是否存储成功
            }
            //  Console.WriteLine("{0}", rows);

            Console.ReadLine();
        }


        public void getColumnDouble(ExcelEdit ed)                                //获取double型的数据
        {

            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //选择指定工作表
                                                                                 //获取选选择的工作表
                                                                                 // Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);//方法一：指定工作表名称读取
                                                                                 //Worksheet ws = (Worksheet)openwb.Worksheets.get_Item(1);//方法二：通过工作表下标读取                                                                              
            int rows = worksheet.UsedRange.Rows.Count;                           //获取工作表中的行数  
            int columns = worksheet.UsedRange.Columns.Count;                     //获取工作表中的列数
            Console.WriteLine("请输入你要获取哪列数据");
            int column = Convert.ToInt16(Console.ReadLine());
            int m = 0;

            //提取对应行列的数据并将其存入数组中，按列读取
            for (int i = 2; i <= rows; i++)
            {
                int temp;
                string a = (worksheet.Cells[i, column]).Text.ToString();
                temp = Convert.ToInt32(a);
                Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX1.Add(temp);
                m++;
            }

            //遍历数组
            foreach (string db in ColumnDB)
            {
                Console.WriteLine("列表中的数据" + db);//查看数组中的数据，测试是否存储成功
            }
            Console.WriteLine("{0}", rows);

            Console.ReadLine();
        }





        public void getTestnum(ExcelEdit ed)                                                 //获取测试集数据
        {
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //选择指定工作表
                                                                                 //获取选选择的工作表
                                                                                 // Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);//方法一：指定工作表名称读取
                                                                                 //Worksheet ws = (Worksheet)openwb.Worksheets.get_Item(1);//方法二：通过工作表下标读取                                                                              
            int rows = worksheet.UsedRange.Rows.Count;                           //获取工作表中的行数  
            int columns = worksheet.UsedRange.Columns.Count;                     //获取工作表中的列数
                                                                                 //  Console.WriteLine("请输入你要获取哪列数据");
            int column = 0;

            //= Convert.ToInt16(Console.ReadLine());
            //提取对应行列的数据并将其存入数组中，按列读取
            //for (int i = 2; i < rows; i++)
            //{
            //    string a = (worksheet.Cells[i, column]).Text.ToString();
            //    Console.WriteLine("读取的数据:" + a);//测试是否获得数据
            //    ColumnDB.Add(a);
            //}
            //按行读取
            for (int i = 2; i <= rows; i++)    //读第一列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 1]).Text.ToString();
                temp = Convert.ToInt32(a);
            //    Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX1.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第二列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 2]).Text.ToString();
                temp = Convert.ToInt32(a);
             //   Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX2.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第3列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 3]).Text.ToString();
                temp = Convert.ToInt32(a);
           //     Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX3.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第4列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 4]).Text.ToString();
                temp = Convert.ToDouble(a);
              //  Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX4.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第5列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 5]).Text.ToString();
                temp = Convert.ToDouble(a);
          //      Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX5.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第6列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 6]).Text.ToString();
                trainY1.Add(Convert.ToDouble(a));
                // temp = Convert.ToDouble(a);
              //  Console.WriteLine("读取的数据:{0}", Convert.ToDouble(a));//测试是否获得数据
                                                                    //     ColumnDB.Add(a)
                                                                    //  trainY.Add(temp+0.0);
                                                                    //   Console.WriteLine("读取的数据1:{0}", trainY[i]);
            }

            for (int i = 2; i <= rows; i++)    //读第7列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 7]).Text.ToString();
                trainY2.Add(Convert.ToDouble(a));
                // temp = Convert.ToDouble(a);
           //     Console.WriteLine("读取的数据:{0}", Convert.ToDouble(a));//测试是否获得数据
                                                                    //     ColumnDB.Add(a)
                                                                    //  trainY.Add(temp+0.0);
                                                                    //   Console.WriteLine("读取的数据1:{0}", trainY[i]);
            }

            for (int i = 2; i <= rows; i++)    //读第8列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 8]).Text.ToString();
                trainY3.Add(Convert.ToDouble(a));
                // temp = Convert.ToDouble(a);
            //    Console.WriteLine("读取的数据:{0}", Convert.ToDouble(a));//测试是否获得数据
                                                                
            }
        }




        

    }

    class ExcelEdit2
    {
        public string mFilename;
        public Microsoft.Office.Interop.Excel.Application app;
        public Microsoft.Office.Interop.Excel.Workbooks wbs;
        public Microsoft.Office.Interop.Excel.Workbook wb;
        public Microsoft.Office.Interop.Excel.Worksheets wss;
        public Microsoft.Office.Interop.Excel.Worksheet ws;
        public List<string> ColumnDB = new List<string>();
        public List<int> trainX1 = new List<int>();
        public List<int> trainX2 = new List<int>();
        public List<int> trainX3 = new List<int>();
        public List<double> trainX4 = new List<double>();
        public List<double> trainX5 = new List<double>();
        public List<double> trainY1 = new List<double>();
        public List<double> trainY2 = new List<double>();
        public List<double> trainY3 = new List<double>();

        public void Open(string FileName)//打开一个Microsoft.Office.Interop.Excel文件
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            wbs = app.Workbooks;
            wb = wbs.Add(FileName);
            mFilename = FileName;
        }
        public Microsoft.Office.Interop.Excel.Worksheet GetSheet(string SheetName)
        //获取一个工作表
        {
            Microsoft.Office.Interop.Excel.Worksheet s = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[SheetName];
            return s;
        }

        public void Close()
        //关闭一个Microsoft.Office.Interop.Excel对象，销毁对象
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
        public bool Save()
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

        public void SetCellValue(Microsoft.Office.Interop.Excel.Worksheet ws, int x, int y, object value)
        //ws：要设值的工作表     X行Y列     value   值
        {
            ws.Cells[x, y] = value;

        }
        public void SetCellValue(string ws, int x, int y, object value)
        //ws：要设值的工作表的名称 X行Y列 value 值
        {

            GetSheet(ws).Cells[x, y] = value;
        }


        public void getTest2num(ExcelEdit2 ed)                                                 //获取测试集数据
        {
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1");  //选择指定工作表
                                                                                 //获取选选择的工作表
                                                                                 // Worksheet ws = ((Worksheet)openwb.Worksheets["Sheet1"]);//方法一：指定工作表名称读取
                                                                                 //Worksheet ws = (Worksheet)openwb.Worksheets.get_Item(1);//方法二：通过工作表下标读取                                                                              
            int rows = worksheet.UsedRange.Rows.Count;                           //获取工作表中的行数  
            int columns = worksheet.UsedRange.Columns.Count;                     //获取工作表中的列数
                                                                                 //  Console.WriteLine("请输入你要获取哪列数据")

            for (int i = 2; i <= rows; i++)    //读第一列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 1]).Text.ToString();
                temp = Convert.ToInt32(a);
             //   Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX1.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第二列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 2]).Text.ToString();
                temp = Convert.ToInt32(a);
             //   Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX2.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第3列数据
            {
                int temp;
                string a = (worksheet.Cells[i, 3]).Text.ToString();
                temp = Convert.ToInt32(a);
              //  Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX3.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第4列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 4]).Text.ToString();
                temp = Convert.ToDouble(a);
              //  Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX4.Add(temp);
            }

            for (int i = 2; i <= rows; i++)    //读第5列数据
            {
                double temp;
                string a = (worksheet.Cells[i, 5]).Text.ToString();
                temp = Convert.ToDouble(a);
            //    Console.WriteLine("读取的数据:{0}", temp);//测试是否获得数据
                ColumnDB.Add(a);
                trainX5.Add(temp);
            }
         
        }
    }


}
