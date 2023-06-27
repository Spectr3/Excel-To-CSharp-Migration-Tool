using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelToModelGenerator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var filePath = "PATH TO EXCEL SHEET";

            Console.WriteLine("Please enter the namespace for the generated classes:");
            var namespaceName = Console.ReadLine();

            Console.WriteLine("Do you want to add decorators to the classes and properties? (yes/no):");
            var addDecorators = Console.ReadLine()?.ToLower() == "yes";

            Console.WriteLine("Please enter the name of the folder to place the class files in:");
            var folderName = Console.ReadLine();

            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }

            var generator = new ClassGenerator();

            var app = new Application();
            var workbook = app.Workbooks.Open(filePath);

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                var originalClassName = worksheet.Name;
                var className = CleanUpClassName(originalClassName);
                var headers = worksheet.Range["1:1"].Cells.Cast<Microsoft.Office.Interop.Excel.Range>()
                    .TakeWhile(cell => cell.Value2 != null)
                    .Select(cell => new { OriginalName = cell.Value2.ToString(), CleanName = CleanUpPropertyName(cell.Value2.ToString()) })
                    .ToArray();

                var classDefinition = generator.GenerateClass(namespaceName, className, headers, addDecorators);
                File.WriteAllText(Path.Combine(folderName, $"{className}.cs"), classDefinition);
            }

            workbook.Close();
            app.Quit();
        }

        private static string CleanUpPropertyName(string propertyName)
        {
            var cleanName = propertyName.Replace(" ", "_")
                .Replace("-", "_")
                .Replace("/", "_");

            if (char.IsDigit(cleanName[0]))
            {
                cleanName = "_" + cleanName;
            }

            return cleanName;
        }

        private static string CleanUpClassName(string className)
        {
            var cleanName = className.Replace(" ", "_")
                .Replace("-", "_")
                .Replace("/", "_");

            if (char.IsDigit(cleanName[0]))
            {
                cleanName = "_" + cleanName;
            }

            return cleanName;
        }
    }
}