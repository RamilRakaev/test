using System;
using System.Collections.Generic;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            Dictionary<int, string> dict = new()
            {
                { 1, "str1" },
                { 2, "str2" },
            };

            dict[10] = "key";
        }

    }
}