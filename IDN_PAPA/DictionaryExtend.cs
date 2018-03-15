using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IDN_PAPA
{
    class DictionaryExtend<TKey, TValue>
    {
       public Dictionary<TKey, TValue> elementDict2 { get; set; }
       public  Dictionary<TKey, TValue> elementsDict1 { get; set; }

       public DictionaryExtend(Dictionary<TKey, TValue> elementDict2, Dictionary<TKey, TValue> elementsDict1)
        {
            this.elementDict2 = elementDict2;
            this.elementsDict1 = elementsDict1;
        }
    }
}
