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
            string[] list1 = sourse1.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            string[] list2 = sourse2.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            Dictionary<string, int> dict1 = IDN.DelRepeats(list1);
            Dictionary<string, int> dict2 = IDN.DelRepeats(list2);
            Dictionary<string, int>[] arrInput = new Dictionary<string, int>[] { dict1, dict2 };
            Dictionary<string, int>[] arrRezult= IDN.SearchInDictionary<string>(dict1, dict2);

            path = SelectFolder();


            WritingInExcMethod(arrInput, arrRezult);
        } 

        private void button2_Click(object sender, EventArgs e)
        {
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt";
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory().ToString();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path = openFileDialog1.FileName;
                sourse1 = File.ReadAllText(path);
                AddToTextBox(textBox1, "Open file:\r\n" + path) ;
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
            string path= Directory.GetCurrentDirectory().ToString();
            folderBrowserDialog1.SelectedPath = path;
            if (folderBrowserDialog1.ShowDialog()==DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
                return path;
            }
            else
            {
                return path;
            }
            
        }

        private void CreateAndWriteToFile(string path, string content)
        {

            FileStream fs= File.Create(path);
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
            foreach(T key in dict.Keys)
            {
                strB.Append(dict[key]+" ");
                strB.Append("");
                strB.Append(key);
                strB.Append("\r\n");
            }
            return strB.ToString();
        }

        private void AddToTextBox(TextBox text_box,string str)
        {
            text_box.Text += str + "\r\n";
        }

        private void WritingInExcMethod<T>(Dictionary<T, int>[] arrInp,Dictionary<T, int>[] arrOut)
        {
            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbooks ObjWorkBooks = null;
            Excel.Workbook ObjWorkBook=null;
            Excel.Worksheet ObjWorkSheet=null;
           

            try
            {
                ObjWorkBooks = ObjExcel.Workbooks;
                ObjWorkBook = ObjWorkBooks.Add(System.Reflection.Missing.Value);
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];


                {
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
                        if (i % 2 == 0) data[0, 0] = "Анализируемый список елементов" ; else data[0, 0] = "Список разыскиваемых елементов";
                        data[1, 0] = "кол-во";
                        data[1, 1] = "елемент";

                        Excel.Range range = ObjWorkSheet.Range[ObjWorkSheet.Cells[1, col], ObjWorkSheet.Cells[1, col + 1]];
                        range.Merge(); //объеденение 2 ячеек
                        //создать диапазон(Range)
                        var startCell = ObjWorkSheet.Cells[1, col];
                        var endCell = ObjWorkSheet.Cells[arrInp[i].Count + vShift, col + 1];
                        var writeRange = ObjWorkSheet.Range[startCell, endCell];


                        //запись данных в диапазон
                        writeRange.Value2 = data;
                        startCell = null;
                        endCell = null;
                        writeRange = null;
                        range = null;
                    }
                }


                {//записали результат-4 новые таблицы
                    int colShift = arrInp.Length * 3;
                    int vShift = 2;
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
                                break;
                            case 1:
                                str = "Не запрошенные аргументы из анализируемого списка аргументов";
                                break;
                            case 2:
                                str = "Аргументы из списка разыскиваемых аргументов НАЙДЕНЫЕ в анализируемом списке аргументов";
                                break;
                            case 3:
                                str = "Аргументы из списка разыскиваемых аргументов НЕ найденые в анализируемом списке аргументов";
                                break;
                            default:
                                str = "";
                                break;

                        }
                        data[0, 0] = str;

                        data[1, 0] = "кол-во";
                        data[1, 1] = "елемент";

                        Excel.Range range = ObjWorkSheet.Range[ObjWorkSheet.Cells[1, col], ObjWorkSheet.Cells[1, col + 1]];
                        range.Merge(); //объеденение 2 ячеек
                        //создать диапазон(Range)
                        var startCell = ObjWorkSheet.Cells[1, col];
                        var endCell = ObjWorkSheet.Cells[arrOut[i].Count + vShift, col + 1];
                        var writeRange = ObjWorkSheet.Range[startCell, endCell];

                        
                        //запись данных в диапазон
                        writeRange.Value2 = data;
                        startCell = null;
                        endCell = null;
                        writeRange = null;
                        range = null;
                    }
                   
                }

                ObjWorkBook.SaveAs(path + @"\rezult.xlsx");//сохранить файл excel

                AddToTextBox(textBox1, "Created file:\r\n" + path + @"\rezult.xlsx");//записать в лог
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
            }
            finally
            {
                ObjExcel.Quit();
                ObjExcel = null;
                ObjWorkBooks = null;
                ObjWorkBook = null;
                ObjWorkSheet = null;
                GC.Collect();
                

            }
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
    }
}
