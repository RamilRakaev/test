using System;
using System.Collections.Generic;
using System.IO;

namespace LinkLibrary
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Console.ReadLine();
            LinksManager manager = new (path);
            manager.OutputLinks = (SourceReferencing[] reference) =>
            {
                foreach (var source in reference)
                {
                    Console.WriteLine($"{source.Titles}\nСсылки:");
                    foreach(var link in source.Links)
                    {
                        Console.WriteLine($"{link}");
                    }
                }
            };
            Console.WriteLine("Hello World!");
        }
    }

    public class SourceReferencing
    {
        public string Titles { get; set; }
        public string[] Links { get; set; }
    }

    public class LinksManager
    {
        public Action<SourceReferencing[]> OutputLinks;
        private readonly string _path;

        public LinksManager(string path)
        {
            _path = path;
        }

        public ReadResult ReadPageLinks()
        {
            if (File.Exists(_path))
            {
                using (StreamReader reader = new (_path))
                {
                    var references = LinksParsing.Parse(reader.ReadToEnd());
                    OutputLinks(references);
                    return ReadResult.Success;
                }
            }
            return ReadResult.FileNotFound;
        }

        public enum ReadResult { FileNotFound, Success }
    }

    internal class LinksParsing
    {
        const string separator = "###";
        const string and = "&&&";

        public static SourceReferencing[] Parse(string page)
        {
            var lines = page.Split("\n", StringSplitOptions.RemoveEmptyEntries);
            var references = new List<SourceReferencing>();
            foreach(var line in lines)
            {
                if (line.Contains("separator"))
                {
                    var sourceReferencing = line.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                    var titles = sourceReferencing[0];
                    var links = sourceReferencing[1];
                    foreach (var title in links.Split(and, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var reference = new SourceReferencing()
                        {
                            Titles = title,
                            Links = titles.Split(and, StringSplitOptions.RemoveEmptyEntries),
                        };
                        references.Add(reference);
                    }
                }
            }
            return references.ToArray();
        }
    }

}
