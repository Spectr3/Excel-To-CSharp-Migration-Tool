using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToModelGenerator
{
    public class ClassGenerator
    {
        public string GenerateClass(string namespaceName, string className, dynamic[] properties, bool addDecorators)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"namespace {namespaceName}");
            builder.AppendLine("{");

            if (addDecorators)
            {
                builder.AppendLine($"\t[WorkbookName(\"{className}\")]");
            }

            builder.AppendLine($"\tpublic class {className}");
            builder.AppendLine("\t{");

            foreach (var property in properties)
            {
                if (addDecorators)
                {
                    builder.AppendLine($"\t\t[WorkbookItemName(\"{property.OriginalName}\")]");
                }

                builder.AppendLine($"\t\tpublic string {property.CleanName} {{ get; set; }} = string.Empty;");
            }

            builder.AppendLine("\t}");
            builder.AppendLine("}");

            return builder.ToString();
        }
    }
}
