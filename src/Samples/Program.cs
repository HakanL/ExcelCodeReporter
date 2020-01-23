using System;
using System.Diagnostics;

namespace Haukcode.ExcelCodeReporter.Samples
{
    public class Program
    {
        public static void Main()
        {
            Console.WriteLine("Samples!");

            var tester = new SimpleExample1();
            string filename = tester.Execute();

            Console.WriteLine($"Output filename: {filename}");

            Process.Start("explorer", filename);
        }
    }
}
