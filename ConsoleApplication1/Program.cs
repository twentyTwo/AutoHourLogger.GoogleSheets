using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            stringSplitter("A12:BB123");
        }

        public static void stringSplitter(string str)
        {
            var strs = str.Split(':').ToArray();
            var first = strs[0];
            var second = strs[1];

            stringProcessor(first);

        }

        public static void stringProcessor(string first, out string str1, out int num1)
        {
            string str = "";
            int num = 0;

            foreach (char f in first)
            {
                if (f >= 'A' && f <= 'Z')
                {
                    str+= f;
                }
                else
                {
                    num = num*10 + (f-'0');
                }

            }

            str1 = str;
            num1 = num;
        }

        

    }

    public class SheetData
    {
        public string StrPart;
        public int NumPart;

    }
}
