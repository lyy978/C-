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

    class Knn
    {
        
        public void distance(ExcelEdit ed, ExcelEdit2 ed2, int len1, int len2, ref List<int> trainX1, ref List<int> trainX2,ref List<int> trainX3,ref List<double> trainX4, ref List<double> trainX5,ref List<double> trainY2, ref List<double> trainY1,ref List<double> trainY3, ref List<int> testX1, ref List<int> testX2,ref List<int> testX3,ref List<double> testX4,ref List<double> testX5, ref List<double> testY1,ref  List<double> testY2,ref List<double> testY3)
        {
           

            List<double> dis = new List<double>();  //所有距离
            double temp = 0;
            double temp1 = 0;
            double temp2 = 0;
            double temp3 = 0;
            double temp4 = 0;
            double temp5 = 0;
            int i = 0;                              //训练集长度 
            int j = 0;                              //测试集长度
            int k = 0;                              //最短距离下标
            double min;
            int count = 0;

           
                    for(j=0;j<testX1.Count;j++)
                    {
                          dis.Clear();
                          Console.WriteLine("discount:{0}", dis.Count); 
                          count++;
                        for (i = 0; i < trainX1.Count; i++)
                        {
                            temp1 = (testX1[j] - trainX1[i]) * (testX1[j] - trainX1[i]);
                            temp2 = (testX2[j] - trainX2[i]) * (testX2[j] - trainX2[i]);
                            temp3 = (testX3[j] - trainX3[i]) * (testX3[j] - trainX3[i]);
                            temp4 = (testX4[j] - trainX4[i]) * (testX4[j] - trainX4[i]);
                            temp5 = (testX5[j] - trainX5[i]) * (testX5[j] - trainX5[i]);
                            temp = temp1 + temp2 + temp3 + temp4 + temp5;
                            dis.Add(temp);
                        }
                    min = (double)dis.Min<double>();

                    Console.WriteLine("最小值:{0}", min);//查看数组中的数据，测试是否存储成功
                    for (i = 1; i < trainX1.Count; i++)
                    {
                      // Console.WriteLine("dis:{0}", dis[i]);
                        if (dis[i]==min)
                        {
                            
                            k = i;
                        break;
                            //Console.WriteLine("dis:{0}", dis[i]);
                            //Console.WriteLine("min:{0}", min );
                            //Console.WriteLine("i:{0}",i);
                        }
                    }
                    Console.WriteLine("dis:{0}", dis[k]);
                    Console.WriteLine("min:{0}", min);
                    Console.WriteLine("k:{0}", k);
                        testY1.Add(trainY1[k]);

                        testY2.Add(trainY2[k]);

                        testY3.Add(trainY3[k]);
                        Console.WriteLine("count:{0}", count);
                     dis.Clear();
                 }
        }

    }
    class Program
    {
        static void Main(string[] args)
        {
            List<int> trainX1 = new List<int>();
            List<int> trainX2 = new List<int>();
            List<int> trainX3 = new List<int>();
            List<double> trainX4 = new List<double>();
            List<double> trainX5 = new List<double>();
            List<double> trainY2 = new List<double>();
            List<double> trainY1 = new List<double>();
            List<double> trainY3 = new List<double>();

            List<int> testX1 = new List<int>();
            List<int> testX2 = new List<int>();
            List<int> testX3 = new List<int>();
            List<double> testX4 = new List<double>();
            List<double> testX5 = new List<double>();
            List<double> testY2 = new List<double>();
            List<double> testY1 = new List<double>();
            List<double> testY3 = new List<double>();

            ExcelEdit ed = new ExcelEdit();              //控制表单的实体
            ed.Open("E:\\C#代码\\knnprogram\\result1.xlsx");      //打开一个excel文件
            Excel.Worksheet worksheet = (Excel.Worksheet)ed.GetSheet("sheet1"); //选择指定的sheet
            ed.getTestnum(ed);                            //获取训练集数据，全部获取   

            ExcelEdit2 ed2 = new ExcelEdit2();              //控制表单的实体
            ed2.Open("E:\\C#代码\\knnprogram\\test.xlsx");      //打开一个excel文件
            Excel.Worksheet worksheet2 = (Excel.Worksheet)ed2.GetSheet("sheet1"); //选择指定的sheet
            ed2.getTest2num(ed2);                           //获取测试集
            Console.WriteLine("行数:{0}", worksheet2.UsedRange.Rows.Count);




            trainX1 = ed.trainX1;
            trainX2 = ed.trainX2;
            trainX3 = ed.trainX3;
            trainX4 = ed.trainX4;
            trainX5 = ed.trainX5;
            trainY1 = ed.trainY1;
            trainY2 = ed.trainY2;
            trainY3 = ed.trainY3;

            testX1 = ed2.trainX1;
            testX2 = ed2.trainX2;
            testX3 = ed2.trainX3;
            testX4 = ed2.trainX4;
            testX5 = ed2.trainX5;
           




            Knn knn = new Knn();
            knn.distance(ed,ed2, worksheet.UsedRange.Rows.Count, worksheet2.UsedRange.Rows.Count, ref trainX1, ref trainX2, ref trainX3, ref trainX4, ref trainX5, ref trainY2,ref trainY1, ref trainY3, ref testX1, ref testX2, ref testX3, ref testX4, ref testX5, ref testY1, ref testY2, ref testY3);
            ed2.trainY1 = testY1;
            ed2.trainY2 = testY2;
            ed2.trainY3 = testY3;

            foreach (double db in ed2.trainY1)
            {
                Console.WriteLine("列表中y1的数据:{0}" ,db);//查看数组中的数据，测试是否存储成功
            }
            foreach (double db in ed2.trainY2)
            {
                Console.WriteLine("列表中y2的数据:{0}" , db);//查看数组中的数据，测试是否存储成功
            }
            foreach (double db in ed2.trainY3)
            {
                Console.WriteLine("列表中y3的数据:{0}" , db);//查看数组中的数据，测试是否存储成功
            }

            for (int x = 0; x < testY1.Count; x++)
            {
                ed2.GetSheet("sheet1").Cells[x + 2, 6] = testY1[x];           //Write the predicted value to sheet1 rowx+2 ,column 6
            }
            for (int x = 0; x < testY1.Count; x++)
            {
                ed2.GetSheet("sheet1").Cells[x + 2, 7] = testY2[x];           //Write the predicted value to sheet1 rowx+2 ,column 6
            }
            for (int x = 0; x < testY1.Count; x++)
            {
                ed2.GetSheet("sheet1").Cells[x + 2, 8] = testY3[x];           //Write the predicted value to sheet1 rowx+2 ,column 6
            }
            ed2.Save();
            ed2.Close();
            Console.ReadLine();
            

        }

       
    }
}
