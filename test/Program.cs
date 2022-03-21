using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace test
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("bla bla 2aa23r".GetTextAfterSubstring("2aa"));
            Console.ReadKey();
        }
    }

    public static class Text
    {
        public static string GetTextAfterSubstring(this string text, string character, int number = 1)
        {
            int index = 0;
            index = text.IndexOf(character) + character.Length;
            if(text.Length == index)
            {
                return string.Empty;
            }
            return text[index..].Trim(' ');
        }
    }
}

