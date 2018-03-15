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

        public List<DictionaryExtend<string, int>> dictMatch { get; set; } = new List<DictionaryExtend<string, int>>();
        public Dictionary<string, int> noMatchElementsArr1 { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, int> noMatchElementsArr2 { get; set; } = new Dictionary<string, int>();

        public Dictionary<string, int> Idn { get; private set; }
        public string[] IdnArr { get; private set; }
        
        public IDN()
        {
          
           
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


         public void SearchInDictionaryCheck(Dictionary<string, int> arr1, Dictionary<string, int> arr2)
        {
            Dictionary<string, int> UsesElementsArr1 = new Dictionary<string, int>();
            Dictionary<string, int> temp1;
            Dictionary<string, int> temp2;
            DictionaryExtend<string, int> tempExtend;

            foreach (string key in arr2.Keys)
            {
                temp1= Search<string>(key, arr1);
                temp2 = new Dictionary<string, int>();
                temp2.Add(key, arr2[key]);
                
                if (temp1.Count!=0)
                {
                    tempExtend = new DictionaryExtend<string, int>(temp2, temp1);
                    dictMatch.Add(tempExtend);

                    foreach(string temp1key in temp1.Keys)
                    {
                        if(!UsesElementsArr1.ContainsKey(temp1key))
                        {
                            UsesElementsArr1.Add(temp1key, temp1[temp1key]);
                        }
                    }
                        
                }
                else
                {
                    noMatchElementsArr2.Add(key, arr2[key]);
                }
            }

            foreach (string key in arr1.Keys)
            {
                if (!UsesElementsArr1.ContainsKey(key))
                {
                    noMatchElementsArr1.Add(key, arr1[key]);
                }
            }

            
        }

        static public Dictionary<T, int> Search<T>(T pattern, Dictionary<T, int> arr1)
        {
            Regex regex = new Regex(pattern.ToString());
            Dictionary<T, int> result = new Dictionary<T, int>();

            foreach (T key in arr1.Keys)
            {
                if (regex.IsMatch(key.ToString()))
                {
                    result.Add(key, arr1[key]);
                }
            }
            return result;
        }
    }
}
