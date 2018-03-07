using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace IDN_PAPA
{
    class IDN
    {
        public Dictionary<string, int> Idn { get; private set; }
        public string[] IdnArr { get; private set; }
        private string pattern;
        public IDN(string[] idns,string pattern)
        {
           idns.CopyTo(IdnArr,0);
           Pattern = pattern;
        }
        static public Dictionary<string, int> DelRepeatsLongTime(string[] arr)
        {
            Stopwatch stop = Stopwatch.StartNew();
            Array.Sort(arr);
            stop.Stop();
            Console.WriteLine(stop.Elapsed);
            Dictionary<string, int> arr2 = new Dictionary<string, int>();
            arr2.Add(arr[0], 1);
            for (int i = 1; i < arr.Length; i++)
            {
                try
                {
                    arr2[arr[i]]++;
                }
                catch
                {
                    arr2.Add(arr[i], 1);
                }
            }
            return arr2;
        }
        static public Dictionary<string, int> DelRepeats(string[] arr)
       {
            Dictionary<string, int> arr2 = new Dictionary<string, int>();
            for(int i=0;i<arr.Length;i++)
            {
                if(!arr2.ContainsKey(arr[i]))
                {
                    arr2.Add(arr[i], 1);
                }
                else
                {
                    arr2[arr[i]]++;
                }
            }
            return arr2;
        }
        static public Dictionary<T, int> DelRepeats<T>(T[] arr)
        {
            Dictionary<T, int> arr2 = new Dictionary<T, int>();
            for (int i = 0; i < arr.Length; i++)
            {
                if (!arr2.ContainsKey(arr[i]))
                {
                    arr2.Add(arr[i], 1);
                }
                else
                {
                    arr2[arr[i]]++;
                }
            }
            return arr2;
        }
        static public Dictionary<T,int>[] SearchInDictionary<T>(Dictionary<T,int> arr1, Dictionary<T, int> arr2)
        {
            Dictionary<T, int> matchElementsArr1 = new Dictionary<T, int>();
            Dictionary<T, int> matchElementsArr2 = new Dictionary<T, int>();
            Dictionary<T, int> noMatchElementsArr1 = new Dictionary<T, int>();
            Dictionary<T, int> noMatchElementsArr2 = new Dictionary<T, int>();
            Dictionary<T, int>[] arrDictionary = new Dictionary<T, int>[] { matchElementsArr1, noMatchElementsArr1, matchElementsArr2, noMatchElementsArr2 };
            

            foreach (T key in arr1.Keys)
            {
                if(arr2.ContainsKey(key))
                {
                    matchElementsArr1.Add(key, arr1[key]);
                    matchElementsArr2.Add(key, arr2[key]);
                }
                else
                {
                    noMatchElementsArr1.Add(key, arr1[key]);
                }
            }

            foreach (T key in arr2.Keys)
            {
                if(!matchElementsArr2.ContainsKey(key))
                {
                    noMatchElementsArr2.Add(key, arr2[key]);
                }
            }
           
            return arrDictionary;

        }
        string Pattern
        {
            get
            {
                return pattern;
            }
            set
            {
                
                try
                {
                    Regex r = new Regex(value);
                    pattern = value;
                }
                catch
                {
                    pattern = @"";
                }
                
            }

        }
    }
}
