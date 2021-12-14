using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace test
{
    interface IA
    {

    }

    class A : IA
    {

    }

    interface IB
    {

    }

    class B
    {

    }

    public class Node
    {
        public Node Parent { get; set; }
        public string Name { get; set; }
        public IEnumerable<Node> Children { get; set; }

        public void GetChildrenNames(in Action<string> action)
        {
            GetChildrenNames(action, this);
        }

        private void GetChildrenNames(in Action<string> action, Node node)
        {
            action($"Наименование {node.Name}. Уровень {node.GetLevel()}");
            if (node.Children == null)
            {
                return;
            }
            foreach (var child in node.Children)
            {
                GetChildrenNames(action, child);
            }
        }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            #region
            Node parent = new();
            parent.Name = "Fucking parent";

            Node child1 = new()
            {
                Name = "child1",
                Parent = parent
            };

            Node child2 = new()
            {
                Name = "child2",
                Parent = child1
            };

            Node child3 = new()
            {
                Name = "child3",
                Parent = child2
            };

            Node child4 = new()
            {
                Name = "child4",
                Parent = child3
            };
            #endregion

            Console.WriteLine((new List<string>() { }).Remove(""));
            //Console.WriteLine(child4.GetLevel());

            #region
            parent.Children = new Node[]
            {
                child1
            };

            child1.Children = new Node[]
            {
                child2, child3
            };

            child2.Children = new Node[]
            {
                child3
            };

            child3.Children = new Node[]
            {
                child4
            };
            #endregion

            //parent.GetChildrenNames((string name) => Console.WriteLine(name));

            A value = new();
            var result = (IA)value;
            var asResult = (IA)value;

            var intValue = 1;
            var intResult = (IComparable) intValue;
        }
    }

    public static class NodeExtensions
    {
        public static int GetLevel(this Node node)
        {
            return GetLevel(node, 1);
        }

        private static int GetLevel(Node node, int level)
        {
            if (node.Parent == null)
            {
                return level;
            }
            return GetLevel(node.Parent, ++level);
        }
    }

}

