using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;


namespace IDN_PAPA
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());

            //Stopwatch stop1 = Stopwatch.StartNew();
            //{
            //    Random rand = new Random();
            //    int length = 2000000;
            //    string ran = "1234567890QWERTYUIOPASDFGHJKLZXCVBNM";
            //    string[] arr = new string[length];
            //    for (int i = 0; i < length; i++)
            //    {
            //        string str="";
            //        for(int j=0;j<20;j++)
            //        {
            //            str += ran[rand.Next(0, 9)].ToString();
            //        }
            //        arr[i] =str ;//+ ran[rand.Next(0, 9)];
            //    }

            //    Stopwatch stop = Stopwatch.StartNew();
            //    Dictionary<string, int> dict = IDN.DelRepeats(arr);
            //    stop.Stop();
            //    Console.WriteLine(stop.Elapsed);
            //}
            //{
            //    Random rand = new Random();
            //    int length = 2000000;
            //    int[] arr = new int[length];
            //    for (int i = 0; i < length; i++)
            //    {
            //        arr[i] = rand.Next(0, 200000);
            //    }

            //    Stopwatch stop = Stopwatch.StartNew();
            //    Dictionary<int, int> dict = IDN.DelRepeats<int>(arr);
            //    stop.Stop();
            //    Console.WriteLine(stop.Elapsed);
            //}


            //{
            //    Random rand = new Random();
            //    int length = 2000000;
            //    string ran = "1234567890QWERTYUIOPASDFGHJKLZXCVBNM";
            //    string[] arr1 = new string[length/3];
            //    string[] arr2 = new string[length];
            //    for (int i = 0; i < length/3; i++)
            //    {
            //        string str = "";
            //        for (int j = 0; j < 20; j++)
            //        {
            //            str += ran[rand.Next(0, 9)].ToString();
            //        }
            //        arr1[i] = str;//+ ran[rand.Next(0, 9)];
            //    }
            //    for (int i = 0; i < length; i++)
            //    {
            //        string str = "";
            //        for (int j = 0; j < 20; j++)
            //        {
            //            str += ran[rand.Next(0, 9)].ToString();
            //        }
            //        arr2[i] = str;//+ ran[rand.Next(0, 9)];
            //    }

            //    Stopwatch stop = Stopwatch.StartNew();
            //    Dictionary<string, int> dict1 = IDN.DelRepeats(arr1);
            //    Dictionary<string, int> dict2 = IDN.DelRepeats(arr2);


            //    IDN.SearchInDictionary<string>(dict1, dict2);
            //    stop.Stop();
            //    Console.WriteLine(stop.Elapsed);
            //}
            //stop1.Stop();
            //Console.WriteLine("General " + stop1.Elapsed);
            //Console.ReadKey();
        }
    }
}
