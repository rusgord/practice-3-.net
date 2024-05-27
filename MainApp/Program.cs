using System;
using System.Reflection;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string wordPluginPath = @"C:\Users\shkverz\source\repos\Practic3.net\WordReportPlugin\bin\Debug\WordReportPlugin.dll";
        string excelPluginPath = @"C:\Users\shkverz\source\repos\Practic3.net\ExcelReportPlugin\bin\Debug\ExcelReportPlugin.dll";

        Console.WriteLine("Select:\n1 - Word\n2 - Excel");
        string choice = Console.ReadLine();

        if (choice == "1")
        {
            Assembly wordAssembly = Assembly.LoadFrom(wordPluginPath);
            dynamic wordPlugin = Activator.CreateInstance(wordAssembly.GetType("WordReportPlugin.WordReport"));

            Console.WriteLine("Select the type of Word report (1 - Text, 2 - Table): ");
            string wordChoice = Console.ReadLine();

            if (wordChoice == "1")
            {
                wordPlugin.GenerateReport("I've already done the coursework, so i only have to write some text).");
            }
            else if (wordChoice == "2")
            {
                string[,] data = new string[,]
                {
                    { "Id", "Name" },
                    { "1", "User 1" },
                    { "2", "User 2" },
                    { "3", "User 3" },
                    { "4", "User 4" },
                    { "5", "User 5" },
                    { "6", "User 6" },
                };
                wordPlugin.GenerateReportWithTable(data);
            }
            else
            {
                Console.WriteLine("Invalid choice.");
            }
        }
        else if (choice == "2")
        {
            Assembly excelAssembly = Assembly.LoadFrom(excelPluginPath);
            dynamic excelPlugin = Activator.CreateInstance(excelAssembly.GetType("ExcelReportPlugin.ExcelReport"));

            string[] data = new string[] { "First", "Second", "Third", "Fourth", "Fifth" };
            excelPlugin.GenerateReport(data);
        }
        else
        {
            Console.WriteLine("Invalid choice.");
        }
    }
}