using System;
using System.Collections;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace test
{
    enum FuckingEnum {}
    public class Enumerable
    {
        public static IEnumerable GetEnumerator()
        {
            for (int i = 0; i < 10; i++)
            {
                yield return i;
            }
        }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            foreach(var i in Enumerable.GetEnumerator())
            {

            }
            Console.ReadLine();
        }
    }
}

