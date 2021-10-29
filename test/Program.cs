using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace test
{
    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(GetCommentsFromDocument(@"C:\Users\Public\Downloads\Анкета IT (4).docx"));
            Console.ReadLine();
        }

        public static string GetCommentsFromDocument(string document)
        {
            List<Answer> answers = new List<Answer>();

            string text = null;

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(document, false))
            {
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                QuestionCategory category = new QuestionCategory();

                foreach (var element in mainPart.Document.Body.ChildElements)
                {
                    var type = element.GetType();
                    if (element.LocalName == "tbl")
                    {
                        category = new QuestionCategory();
                        ParseTable(element);
                    }
                }
            }
            return text;
        }

        private static void ParseTable(OpenXmlElement table)
        {
            if (table.LocalName == "tbl")
            {
                foreach (var child in table.ChildElements.Where(e => e.LocalName == "tr"))
                {
                    RowParse(child);
                }
            }
        }

        private static void RowParse(OpenXmlElement tableRow)
        {
            Answer answer = new Answer();
            if (tableRow.LocalName == "tr")
            {
                foreach (var cell in tableRow.ChildElements.Where(e => e.LocalName == "tc"))
                {
                    cell.InnerText
                }
            }
        }
    }
}

