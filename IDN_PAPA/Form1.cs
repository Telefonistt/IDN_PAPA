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
            Dictionary<string, int>[] arr= IDN.SearchInDictionary<string>(dict1, dict2);
            
            string[] contents=new string[arr.Length];
            for(int i=0;i< contents.Length;i++)
            {
                contents[i] = CreateStringList<string>(arr[i]);
            }

            path = SelectFolder();

            
            CreateAndWriteToFile(path + @"\Match_list1.txt", contents[0]);
            AddToTextBox(textBox1, "Created file:\r\n" + path + @"\Match_list1.txt");
            CreateAndWriteToFile(path + @"\NoMatch_list1.txt", contents[1]);
            AddToTextBox(textBox1, "Created file:\r\n" + path + @"\NoMatch_list1.txt");
            CreateAndWriteToFile(path + @"\Match_list2.txt", contents[2]);
            AddToTextBox(textBox2, "Created file:\r\n" + path + @"\Match_list2.txt");
            CreateAndWriteToFile(path + @"\NoMatch_list2.txt", contents[3]);
            AddToTextBox(textBox2, "Created file:\r\n" + path + @"\NoMatch_list2.txt");

            //WritingInExcMethod();
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

        private void WritingInExcMethod()
        {
            try
            {
                Excel.Application ObjExcel = new Excel.Application();
                Excel.Workbook ObjWorkBook;
                Excel.Worksheet ObjWorkSheet;
                ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
                ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
                ObjWorkSheet.Cells[3, 1] = "51";
                ObjWorkBook.SaveAs(path+@"\log.xlsx");
                /**/

                ObjExcel.Quit();
            }
            catch (Exception exc)
            {
                MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
            }
        }
    }
}
