
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Collector
{
    class Program
    {
        static Model model = new Model();
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.InputEncoding = Encoding.GetEncoding("koi8-u");
            Console.Title = "Collector";
            Console.Write("Enter the path to the document: ");
            model.SetPath(Console.ReadLine());
            model.LoadDoc();
            model.LoadTable();
            model.ScanTable();
            Console.ReadKey();
        }
    }
}
