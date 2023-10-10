using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace ConsoleApp1
{
    class Class1
    {
        WebClient wc = new WebClient();

        const string FilePath = @"C:\\test\\Тестовый_файл.xlsx";
        private const string TestFilePath = @"C:\\test\\Тестовый_файл1.xlsx";
        public void Sqve(string message, long? userID)
        {
            try
            {
                string message1 = "";
                long? userID1 = 0;
                bool b = false;
                string txt = "";
                StreamReader fin = new StreamReader(@"C://test1//Name.txt");
                string jo = fin.ReadLine();
                fin.Close();
                if (jo == "" || jo == null)
                {
                    txt = $"{message},{userID}@";
                    b = true;
                }
                else//Вызвано исключение: "System.InvalidCastException" в ConsoleApp3.exe
                {
                    string[] words = jo.Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                    
                    for (int i = 0; i < words.Length; i++)
                    {
                        try
                        {
                            string[] words1 = words[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            if (words1.Length == 2)
                            {
                                message1 = words1[0];
                                userID1 = Convert.ToInt32(words1[1]);
                                if (userID1 == userID && message != null && message != null)
                                {
                                    message1 = message;
                                    b = true;
                                }
                                txt += $"{message1},{userID1}@";
                            }
                        }
                        catch { }
                    }
                }
                StreamWriter f = new StreamWriter(@"C://test1//Name.txt");
                if (b == false)
                {
                    txt += $"{message},{userID}@";
                }

                f.WriteLine(txt);
                f.Close();

                Save.Safe = false;
            }
            catch { }
        }
        public async void asincVoid()
        {
            
            await Task.Run(() => file1());
        }
        public void file1()
        {
            try
            {

                var url = "http://serp-koll.ru/images/ep/k1/rasp1.xlsx";
                if (!System.IO.File.Exists(@"C://test//Тестовый_файл.xlsx"))
                {
                    wc.DownloadFile(url, FilePath);
                }
                wc.DownloadFile(url, TestFilePath);




                FileInfo file = new FileInfo(@"C://test//Тестовый_файл.xlsx");
                long size = file.Length;
                FileInfo srs = new FileInfo(@"C:\\test\\Тестовый_файл1.xlsx");
                long size1 = srs.Length;
                string[,] list1;
                if (size != size1)
                {
                    list1 = Shet(@"C:\\test\\Тестовый_файл1.xlsx");
                    System.IO.File.Delete(FilePath);
                    var filesToRename = Directory.GetFiles(@"C:\\test\\").Select(f => new FileInfo(f));
                    foreach (FileInfo file1 in filesToRename)
                    {
                        File.Move(file1.FullName, @"C:\\test\\Тестовый_файл.xlsx");
                    }
                    Class2.Mas123 = list1;

                    
                    Save.Safe1 = true;
                    string message = "";
                    long? userID = 0;
                    StreamReader fin = new StreamReader(@"C://test1//Name.txt");
                    string jo = fin.ReadLine();
                    if (jo != "" && jo != null)
                    {
                        string[] words = jo.Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (message == null || message == null)
                            {
                                break;
                            }
                            string[] words1 = words[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            message = words1[0];
                            userID = Convert.ToInt32(words1[1]);
                            Console.WriteLine(message);
                            Program.e(message, userID);
                        }
                    }
                    fin.Close();
                    Save.Safe1 = false;



                }
                else
                {
                    File.Delete(@"C:\\test\\Тестовый_файл1.xlsx");
                    if (Class2.Mas123 == null)
                    {
                        Class2.Mas123 = Shet(@"C:\\test\\Тестовый_файл.xlsx");
                    }
                }



                //    string[,] list = Shet(@"C:\\test\\Тестовый_файл.xlsx");
                //string[,] list1 = Shet(@"C:\\test\\Тестовый_файл1.xlsx");
                //var isFileChanged = false;
                //for (int i = 1; i < Class2.collumn; i++) //по всем колонкам
                //{
                //    for (int j = 1; j < Class2.Row; j++) // по всем строкам
                //    {
                //        if (isFileChanged)
                //        {
                //            break;
                //        }

                //        if (list[i, j] != list1[i, j])
                //        {
                //            if (isFileChanged)
                //            {
                //                break;
                //            }
                //            System.IO.File.Delete(FilePath);
                //            var filesToRename = Directory.GetFiles(@"C:\\test\\").Select(f => new FileInfo(f));
                //            foreach (FileInfo file1 in filesToRename)
                //            {
                //                File.Move(file1.FullName, @"C:\\test\\Тестовый_файл.xlsx");
                //            }
                //            Class2.Mas123 = list1;


                //            isFileChanged = true;
                //        }
                //    }

                //}
                //if (!isFileChanged)
                //{
                //    File.Delete(@"C:\\test\\Тестовый_файл1.xlsx");
                //    Class2.Mas123 = list;
                //}
                //else
                //{
                //    Save.Safe1 = true;
                //    string message = "";
                //    long? userID = 0;
                //    StreamReader fin = new StreamReader(@"C://test1//Name.txt");
                //    string jo = fin.ReadLine();
                //    string[] words = jo.Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                //    for (int i = 0; i < words.Length; i++)
                //    {
                //        string[] words1 = words[i].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                //        message = words1[0];
                //        userID = Convert.ToInt32(words1[1]);
                //        Program.e(message, userID);
                //    }
                //    fin.Close();
                //    Save.Safe1 = false;
                //}
                
                try
                {
                File.Delete(@"C:\\test\\Тестовый_файл1.xlsx");
                }
            catch { }
            }
            catch { }
            Thread.Sleep(900000);
            file1();

        }
        public string[,] Shet(string sus)
        {
            try
            {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(sus, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            DateTime date = DateTime.Now;
            int dayofweek = (int)date.DayOfWeek;
            if (dayofweek >= 6) dayofweek = 1;
            Excel.Worksheet ObjWorkSheet;
            try
            {
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[dayofweek];//получить 1 лист
            }
            catch
            {
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];//получить 1 лист
            }

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            int row = lastCell.Row;
            int collumn = lastCell.Column;
            Class2.Row = lastCell.Row;
            Class2.collumn = collumn;
            if (Class2.Row > 50)
                Class2.Row = 50;
            if (Class2.collumn > 35)
                Class2.collumn = 35;




            string[,] list = new string[Class2.collumn, Class2.Row]; // массив значений с листа равен по размеру листу

            for (int i = 1; i < Class2.collumn; i++) //по всем колонкам
                for (int j = 1; j < Class2.Row; j++) // по всем строкам
                    list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку
            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            
            try
            {
                ObjWorkBook.Close();
            }
            catch
            { }
            return list;
            }
            catch
            {
                return null;
            }
        }
        public void filel()
        {
            try
            {
                Fam();

            }
            catch
            {
            }
        }

        public void file()
        {
            try
            {
                    foo();
            }
            catch
            {
            }
        }
        public void Fam()
        {
            try
            {

                string[,] list = Class2.Mas123;
                string[] mas = new string[7];
                int y = 1;
                int[] gg = new int[6];
                int r = 0;
                string hh = Classr.fam;
                bool u = false;
                for (int i = 0; i < Class2.Row; i++) //по всем колонкам
                {
                    for (int j = 0; j < Class2.collumn/*-1*/; j++) // по всем строкам 
                    {
                        if (list[j, i] != null)
                            u = list[j, i].ToLower().Contains(hh);
                        if (u)
                        {
                            gg[r] = j;
                            r++;
                            mas[y] = list[j, i];
                            y++;
                        }
                        string lol = list[j, i];
                        if (lol != null)
                        {
                            if (lol.StartsWith("Расписание"))
                            {
                                mas[0] = lol;
                            }
                        }

                    }
                }
                Class2.Mas = mas;
                jojo(gg);
            }
            catch { }
            //for (int i = 0; i < 14; i++)
            //{
            //    Console.WriteLine(mas[i]);
            //}
        }
        public void jojo(int[] gg)
        {
            try
            {
                string[] mas = Class2.Mas;
                string[,] list = Class2.Mas123;
                int s;
                int num;
                bool isNum;

                for (int i = 0; i < 6; i++) //по всем колонкам
                {
                    s = gg[i];
                    for (int j = 0; j < Class2.Row/*-1*/; j++)
                    {
                        isNum = int.TryParse(list[s, j], out num);
                        if (isNum)
                        {
                            mas[i + 1] = mas[i + 1] + $" [{list[s, j]}]";
                            break;
                        }

                    }

                }
                Class2.Mas = mas;
            }
            catch { }
        }
        public void foo()
        {
            try
            {
                string[,] list = Class2.Mas123;

                int r = 0;
                string[] mas = new string[14];
                int y = 1;
                int y1 = 0;
                for (int i = 0; i < Class2.collumn; i++) //по всем колонкам
                {
                    for (int j = 0; j < Class2.Row; j++) // по всем строкам 
                    {



                        if (list[i, j] == Class2.Gug)
                        {
                            r = i;
                        }
                        if (i == r)
                        {
                            if (list[i, j] != null)
                            {
                                //Console.Write($"|{list[i, j]}  {i}   {j}  |\n");
                                if (y1 == 0 || y1 == 1 || y1 == 3 || y1 == 11 || y1 == 15 || y1 == 19 || y1 == 23 || y1 == 7 || y1 == 5 || y1 == 9 || y1 == 13 || y1 == 17 || y1 == 21)
                                {

                                    mas[y] = list[i, j];
                                    y++;
                                }
                                y1++;
                            }
                        }
                        string lol = list[i, j];
                        if (lol != null)
                        {
                            if (lol.StartsWith("Расписание"))
                            {

                                mas[0] = lol;
                            }
                        }

                    }
                }
                Class2.Mas = mas;
            }
            catch
            {

            }


        }
    }
}
