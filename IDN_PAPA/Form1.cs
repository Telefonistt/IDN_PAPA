using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace IDN_PAPA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private string sourse1;
        private string sourse2;


        private string path;
        private void button1_Click(object sender, EventArgs e)
        {
            if (sourse1 != null && sourse2 != null)
            {


                string[] list1 = sourse1.Split(new char[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                string[] list2 = sourse2.Split(new char[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, int> dict1 = IDN.DelRepeats(list1);
                Dictionary<string, int> dict2 = IDN.DelRepeats(list2);
                Dictionary<string, int>[] arrInput = new Dictionary<string, int>[] { dict1, dict2 };
                if (checkBox1.Checked == true)
                {

                    Dictionary<string, int>[] arrRezult = IDN.SearchInDictionary<string>(dict1, dict2);

                    path = SelectFolder();
                    if (path == null) return;

                    WritingInExcMethod(arrInput, arrRezult);
                }
                else
                {
                    IDN rezult = new IDN();
                    rezult.SearchInDictionaryCheck(dict1, dict2);

                    List<DictionaryExtend<string, int>> dictMatch = rezult.dictMatch;
                    Dictionary<string, int> noMatchElementsArr1 = rezult.noMatchElementsArr1;
                    Dictionary<string, int> noMatchElementsArr2 = rezult.noMatchElementsArr2;

                    path = SelectFolder();
                    //if (path == null) return;
                    WritingInExcMethod2(arrInput,dictMatch, noMatchElementsArr1, noMatchElementsArr2);



                }

}
            else
            {
                MessageBox.Show("Ошибка, нет информации");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                sourse1 = File.ReadAllText(path);
                AddToTextBox(textBox1, "Open file:\r\n" + path);

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog2.FileName;
                sourse2 = File.ReadAllText(path);
                AddToTextBox(textBox2, "Open file:\r\n" + path);
            }
        }

        private string SelectFolder()
        {
            string path = Directory.GetCurrentDirectory().ToString();
            folderBrowserDialog1.SelectedPath = path;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
                return path;
            }
            else
            {
                return null;
            }

        }

        private void CreateAndWriteToFile(string path, string content)
        {

            FileStream fs = File.Create(path);
            fs.Close();
            //FileStream fs= File.Create(path);
            // fs.Close();
            // string appendText="";//= "This is extra text" + Environment.NewLine;
            // foreach (string line in readText)
            // {
            //     appendText += line + Environment.NewLine;
            // }
            File.AppendAllText(path, content);
        }

        private string CreateStringList<T>(Dictionary<T, int> dict)
        {
            StringBuilder strB = new StringBuilder();
            foreach (T key in dict.Keys)
            {
                strB.Append(dict[key] + " ");
                strB.Append("");
                strB.Append(key);
                strB.Append("\r\n");
            }
            return strB.ToString();
        }

        private void AddToTextBox(TextBox text_box, string str)
        {
            text_box.Text += str + "\r\n";
        }

        private void WritingInExcMethod<T>(Dictionary<T, int>[] arrInp, Dictionary<T, int>[] arrOut)
        {
            Excel.Application ObjExcel;
            bool flagexcelapp = false;
            try
            {// Присоединение к открытому приложению Excel (если оно открыто)
                ObjExcel = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                flagexcelapp = true; // устанавливаем флаг в 1, будем знать что присоединились
            }
            catch
            {
                ObjExcel = new Excel.Application();// Если нет, то создаём новое приложение
            }



            Excel.Workbooks ObjWorkBooks = null;
            Excel.Workbook ObjWorkBook = null;
            Excel.Worksheet ObjWorkSheet1 = null;
            Excel.Worksheet ObjWorkSheet2 = null;



            
                ObjWorkBooks = ObjExcel.Workbooks;
                ObjWorkBook = ObjWorkBooks.Add(System.Reflection.Missing.Value);
                ObjWorkSheet1 = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                int k = ObjWorkBook.Sheets.Count;
                ObjWorkBook.Sheets.Add(After: ObjWorkBook.Sheets.Add(After: (Excel.Worksheet)ObjWorkBook.Sheets[k]), Count: (7 - k));
                ObjWorkSheet1.Name = "Общее";

                
                   int colShift = 0;
                    int vShift = 2;
                    for (int i = 0; i < arrInp.Length; i++)//проход по массиву dictionary
                    {
                        var dict1 = arrInp[i];
                        int row = 1, col = 3 * i + 1 + colShift;
                        var data = new object[arrInp[i].Count + vShift, 2];
                        foreach (T key in dict1.Keys)
                        {
                            data[row - 1 + vShift, 0] = dict1[key].ToString();
                            data[row - 1 + vShift, 1] = key.ToString();
                            row++;
                        }
                        //заголовки
                        if (i % 2 == 0)
                        {
                            data[0, 0] = "Анализируемый список елементов";
                            ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[2];
                            ObjWorkSheet2.Name = "\"Где\"";
                        }
                        else
                        {
                            data[0, 0] = "Список разыскиваемых елементов";
                            ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[3];
                            ObjWorkSheet2.Name = "\"Что\"";
                        }
                        data[1, 0] = "кол-во";
                        data[1, 1] = "елемент";

                        Excel.Range range1 = ObjWorkSheet1.Range[ObjWorkSheet1.Cells[1, col], ObjWorkSheet1.Cells[1, col + 1]];
                        Excel.Range range2 = ObjWorkSheet2.Range[ObjWorkSheet2.Cells[1, 1], ObjWorkSheet2.Cells[1, 2]];
                        range1.Merge(); //объеденение 2 ячеек
                        range2.Merge();
                        //создать диапазон(Range)
                        var startCell1 = (Excel.Range)ObjWorkSheet1.Cells[1, col];
                        var endCell1 = (Excel.Range)ObjWorkSheet1.Cells[arrInp[i].Count + vShift, col + 1];
                        Excel.Range writeRange1 = ObjWorkSheet1.Range[startCell1, endCell1];

                        var startCell2 = (Excel.Range)ObjWorkSheet2.Cells[1, 1];
                        var endCell2 = (Excel.Range)ObjWorkSheet2.Cells[arrInp[i].Count + vShift, 2];
                        Excel.Range writeRange2 = ObjWorkSheet2.Range[startCell2, endCell2];

                        //запись данных в диапазон
                        writeRange1.Value2 = data;
                        writeRange2.Value2 = data;

                        
                    }
                


                //записали результат-4 новые таблицы
                     colShift = arrInp.Length * 3;
                     vShift = 2;
                    for (int i = 0; i < arrOut.Length; i++)//проход по массиву dictionary
                    {
                        var dict1 = arrOut[i];
                        int row = 1, col = 3 * i + 1 + colShift;
                        var data = new object[arrOut[i].Count + vShift, 2];
                        foreach (T key in dict1.Keys)
                        {
                            data[row - 1 + vShift, 0] = dict1[key].ToString();
                            data[row - 1 + vShift, 1] = key.ToString();
                            row++;
                        }
                        //заголовки
                        string str;
                        switch (i)
                        {
                            case 0:
                                str = "Аргументы из списка разыскиваемых аргументов найденые в анализируемом списке аргументов";
                                ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[4];
                                ObjWorkSheet2.Name = "\"Где\" совпад";
                                break;
                            case 1:
                                str = "Не запрошенные аргументы из анализируемого списка аргументов";
                                ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[5];
                                ObjWorkSheet2.Name = "\"Где\" не совпад";
                                break;
                            case 2:
                                str = "Аргументы из списка разыскиваемых аргументов НАЙДЕНЫЕ в анализируемом списке аргументов";
                                ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[6];
                                ObjWorkSheet2.Name = "\"Что\" совпад";
                                break;
                            case 3:
                                str = "Аргументы из списка разыскиваемых аргументов НЕ найденые в анализируемом списке аргументов";
                                ObjWorkSheet2 = (Excel.Worksheet)ObjWorkBook.Sheets[7];
                                ObjWorkSheet2.Name = "\"Что\" не совпад";
                                break;
                            default:
                                str = "";
                                break;

                        }
                        data[0, 0] = str;

                        data[1, 0] = "кол-во";
                        data[1, 1] = "елемент";
                        //объеденение 2 ячеек
                        Excel.Range range1 = ObjWorkSheet1.Range[ObjWorkSheet1.Cells[1, col], ObjWorkSheet1.Cells[1, col + 1]];
                        Excel.Range range2 = ObjWorkSheet2.Range[ObjWorkSheet2.Cells[1, 1], ObjWorkSheet2.Cells[1, 2]];
                        range1.Merge(); //объеденение 2 ячеек
                        //создать диапазон(Range)
                        var startCell1 = (Excel.Range)ObjWorkSheet1.Cells[1, col];
                        var endCell1 = (Excel.Range)ObjWorkSheet1.Cells[arrOut[i].Count + vShift, col + 1];
                        Excel.Range writeRange1 = ObjWorkSheet1.Range[startCell1, endCell1];

                        var startCell2 = (Excel.Range)ObjWorkSheet2.Cells[1, 1];
                        var endCell2 = (Excel.Range)ObjWorkSheet2.Cells[arrOut[i].Count + vShift, 2];
                        Excel.Range writeRange2 = ObjWorkSheet2.Range[startCell2, endCell2];

                        //запись данных в диапазон
                        writeRange1.Value2 = data;
                        writeRange2.Value2 = data;

                        
                    }

                

                ObjWorkBook.SaveAs(path + @"\rezult.xlsx");//сохранить файл excel
               



                // далее закрытие
                // если не присоединялись, а создавали своё приложение то тупо убиваем процессы


                AddToTextBox(textBox1, "Created file:\r\n" + path + @"\rezult.xlsx");//записать в лог

            //catch (Exception exc)
            //{
            //    MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
            //}

            // далее закрытие
            // если не присоединялись, а создавали своё приложение то тупо убиваем процессы
            if (flagexcelapp == false)
            {
                ObjWorkBooks.Close();
                ObjExcel.Quit();
                System.Diagnostics.Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p2 in ps2)
                {
                    p2.Kill();
                }
            }
            else // Если же мы присоединялись, то закрываем рабочую книгу, по поводу параметров
            {    // "false" - можете почитать на MSDN - мне разбираться было лень :)
                ObjWorkBook.Close(false, false, false);
            }


            GC.Collect();
                
        }

        private void WriteToTxt<T>(Dictionary<T, int>[] arr)
        {
            string[] contents = new string[arr.Length];
            for (int i = 0; i < contents.Length; i++)
            {
                contents[i] = CreateStringList<T>(arr[i]);
            }

            CreateAndWriteToFile(path + @"\Match_list1.txt", contents[0]);
            AddToTextBox(textBox1, "Created file:\r\n" + path + @"\Match_list1.txt");
            CreateAndWriteToFile(path + @"\NoMatch_list1.txt", contents[1]);
            AddToTextBox(textBox1, "Created file:\r\n" + path + @"\NoMatch_list1.txt");
            CreateAndWriteToFile(path + @"\Match_list2.txt", contents[2]);
            AddToTextBox(textBox2, "Created file:\r\n" + path + @"\Match_list2.txt");
            CreateAndWriteToFile(path + @"\NoMatch_list2.txt", contents[3]);
            AddToTextBox(textBox2, "Created file:\r\n" + path + @"\NoMatch_list2.txt");
        }




        private void WritingInExcMethod2(Dictionary<string, int>[] arrInp,List<DictionaryExtend<string, int>> dictMatch,Dictionary<string, int> noMatchElementsArr1 ,Dictionary<string, int> noMatchElementsArr2 )
        {
            Excel.Application ExcelApp;
            Excel.Workbooks workBooks;
            Excel.Workbook workBook;
            Excel.Worksheet workSheet;
            Excel.Range workRange;
           

            bool flagexcelapp = false;
            try
            {// Присоединение к открытому приложению Excel (если оно открыто)
                ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                flagexcelapp = true; // устанавливаем флаг в 1, будем знать что присоединились
            }
            catch
            {
                ExcelApp = new Excel.Application();// Если нет, то создаём новое приложение
            }

            


                //ExcelApp.Visible = true;
                workBooks = ExcelApp.Workbooks;
                workBook = workBooks.Add();

                int k = workBook.Sheets.Count;
                if (k <= 2)
                    workBook.Sheets.Add(After: workBook.Sheets.Add(After: (Excel.Worksheet)workBook.Sheets[k]), Count: (3 - k));

            int rowShift = 0;
            int colShift = 0;
    //////////////////////////////////////////////////////2222222222222222222222222///////////////////////////////////////////////
                workSheet = workBook.Sheets[2];
                workSheet.Name = "Совпадения";
                object[,] data = CreateData1(dictMatch);
                int i = data.GetLength(0);
                int j = data.GetLength(1);
                rowShift = 2;
                for(int n=0;n<dictMatch.Count*5;n+=5)
                {
                    workSheet.Cells[1, 1+ n] = "Что искали";
                    workSheet.Cells[2, 1 + n] = "Елемент";
                    workSheet.Cells[2, 2 + n] = "кол-во";
                    workSheet.Cells[1, 3+ n] = "Что нашли по запросу";
                    workSheet.Cells[2, 3 + n] = "Елемент";
                    workSheet.Cells[2, 4 + n] = "кол-во";
                }
                
            var startCell = (Excel.Range)workSheet.Cells[1+ rowShift, 1];
                var endCell = (Excel.Range)workSheet.Cells[i+ rowShift, j];
                
                workRange = workSheet.Range[startCell, endCell];
                //workRange = workSheet.Range[1, 1];

                workRange.Value = data;
                ////////////////////////////////////////////////////22222222222222//////////////////////////////////////////////////////////////////

                workSheet = workBook.Sheets[3];
                workSheet.Name = "Исключения";
                data = CreateData(noMatchElementsArr1);

                 i = data.GetLength(0);
                 j = data.GetLength(1);
                 colShift = 0;
                 workSheet.Cells[1, 1] = "НЕ искали";
                workSheet.Cells[2, 1 + colShift] = "Елемент";
                workSheet.Cells[2, 2 + colShift] = "кол-во";

            if (i != 0)
                {
                    startCell = (Excel.Range)workSheet.Cells[1+rowShift, 1 + colShift];
                    endCell = (Excel.Range)workSheet.Cells[i+rowShift, j + colShift];
                    workRange = workSheet.Range[startCell, endCell];
                    workRange.Value = data;
                }
            ///////////////////////////////////////////////////////////////////////////
                workSheet = workBook.Sheets[3];
                data = CreateData(noMatchElementsArr2);
                colShift = 4;
                i = data.GetLength(0);
                j = data.GetLength(1);
                workSheet.Cells[1, 1+colShift] = "НЕ нашли";
                workSheet.Cells[2, 1 + colShift] = "Елемент";
                workSheet.Cells[2, 2 + colShift] = "кол-во";
            if (i!=0)
                {
                    startCell = (Excel.Range)workSheet.Cells[1+ rowShift, 1 + colShift];
                    endCell = (Excel.Range)workSheet.Cells[i+ rowShift, j + colShift];
                    workRange = workSheet.Range[startCell, endCell];
                    workRange.Value = data;
                }

            ///////////////////////////////////////////////////////////////////////////////////////////////////
            workSheet = workBook.Sheets[1];
            workSheet.Name = "Вход данные";
            data = CreateData(arrInp[0]);
             colShift = 0;
            i = data.GetLength(0);
            j = data.GetLength(1);
            workSheet.Cells[1, 1 + colShift] = "Где ищем";
            workSheet.Cells[2, 1 + colShift] = "Елемент";
            workSheet.Cells[2, 2 + colShift] = "кол-во";

            if (i != 0)
            {
                startCell = (Excel.Range)workSheet.Cells[1+ rowShift, 1 + colShift];
                endCell = (Excel.Range)workSheet.Cells[i+ rowShift, j + colShift];
                workRange = workSheet.Range[startCell, endCell];
                workRange.Value = data;
            }
            ////////////////////////////////////////////////////////////////////////////
            workSheet = workBook.Sheets[1];
            data = CreateData(arrInp[1]);
            colShift = 3;
            i = data.GetLength(0);
            j = data.GetLength(1);
            workSheet.Cells[1, 1+colShift] = "Что ищем";
            workSheet.Cells[2, 1 + colShift] = "Елемент";
            workSheet.Cells[2, 2 + colShift] = "кол-во";
            if (i != 0)
            {
                startCell = (Excel.Range)workSheet.Cells[1+rowShift, 1 + colShift];
                endCell = (Excel.Range)workSheet.Cells[i+ rowShift, j + colShift];
                workRange = workSheet.Range[startCell, endCell];
                workRange.Value = data;
            }
            ///////////////////////////////////////            ////////////////////////////////////////////////////////////////////////////////
            workBook.SaveAs(path + @"\rezult.xlsx");//сохранить файл excel



            if (flagexcelapp == false)
            {
                workBooks.Close();
                ExcelApp.Quit();
                System.Diagnostics.Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                foreach (System.Diagnostics.Process p2 in ps2)
                {
                    p2.Kill();
                }
            }
            else // Если же мы присоединялись, то закрываем рабочую книгу, по поводу параметров
            {    // "false" - можете почитать на MSDN - мне разбираться было лень :)
                workBook.Close(false, false, false);
            }

            
                GC.Collect();
           
            
            



        }

        object[,] CreateData(List<DictionaryExtend<string, int>> dictMatch)
        {
            int length = 0; ;
            int n = dictMatch.Count;
            for (int k = 0; k < n; k++)
            {
                var dictKMatch = dictMatch[k].elementsDict1;
                length += dictKMatch.Count;
            }
            object[,] result = new object[length, 4];
            int index = 0;
            for (int k=0;k<n;k++)
            {
                var dict2Elem = dictMatch[k].elementDict2;
                var dict1Match = dictMatch[k].elementsDict1;

                result[index, 0] = dict2Elem.FirstOrDefault().Key;
                result[index, 1] = dict2Elem.FirstOrDefault().Value;
                //result[index, 2] = dict1Match.FirstOrDefault().Key;
                //result[index, 3] = dict1Match.FirstOrDefault().Value;

                for(int i=index+1;i<dict1Match.Count;i++)
                {
                    result[i, 0] = "";
                    result[i, 1] = "";
                }

                foreach(string key in dict1Match.Keys)
                {
                    result[index, 2] = key;
                    result[index, 3] = dict1Match[key];
                    index++;
                }

            }
            

            //for(int i=0;i<result.GetLength(0);i++)
            //{
            //    for (int j = 0; j < result.GetLength(1); j++)
            //    {
            //        Console.Write(result[i,j] + " ");
            //    }
            //    Console.WriteLine();
            //}

            return result;
        }


        object[,] CreateData1(List<DictionaryExtend<string, int>> dictMatch)
        {
            int maxHeight = 0;
            int maxWidth = dictMatch.Count*5;
            int n = dictMatch.Count;
            for (int k = 0; k < n; k++)
            {
                var dictKMatch = dictMatch[k].elementsDict1;
                if(maxHeight< dictKMatch.Count)
                maxHeight = dictKMatch.Count;
            }
            object[,] result = new object[maxHeight, maxWidth];
            int index = 0;
            int col = 0;
            for (int k = 0; k < n; k++)
            {
                var dict2Elem = dictMatch[k].elementDict2;
                var dict1Match = dictMatch[k].elementsDict1;

                result[index, 0+col] = dict2Elem.FirstOrDefault().Key;
                result[index, 1+ col] = dict2Elem.FirstOrDefault().Value;
                //result[index, 2] = dict1Match.FirstOrDefault().Key;
                //result[index, 3] = dict1Match.FirstOrDefault().Value;

                for (int i = index + 1; i < dict1Match.Count; i++)
                {
                    result[i, 0+ col] = "";
                    result[i, 1+ col] = "";
                }

                foreach (string key in dict1Match.Keys)
                {
                    result[index, 2+ col] = key;
                    result[index, 3+ col] = dict1Match[key];
                    index++;
                }
                index = 0;
                col += 5;
            }


            //for(int i=0;i<result.GetLength(0);i++)
            //{
            //    for (int j = 0; j < result.GetLength(1); j++)
            //    {
            //        Console.Write(result[i,j] + " ");
            //    }
            //    Console.WriteLine();
            //}

            return result;
        }

        object[,] CreateData(Dictionary<string, int> dict)
        {
            int row = 0;
            object[,] result = new object[dict.Count, 2];
            foreach (string key in dict.Keys)
            {
                result[row, 0] = key ;
                result[row, 1] = dict[key];
                row++;
            }
            return result;
        }
    }
}
