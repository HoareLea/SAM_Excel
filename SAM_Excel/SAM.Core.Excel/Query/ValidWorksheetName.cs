using NetOffice.ExcelApi;
using System.Collections.Generic;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static string ValidWorksheetName(this Workbook workbook, string name, string defaultName = "Sheet", int startIndex = 1, string separator = "")
        {
            string result = name;

            if(startIndex.ToString().Length > 31)
            {
                startIndex = 1;
            }

            if(separator == null || separator.Length > 31)
            {
                separator = string.Empty;
            }

            if(string.IsNullOrWhiteSpace(result))
            {
                result = defaultName;
            }
            if (result.StartsWith("'"))
            {
                result = result.Substring(1);
            }

            if (result.EndsWith("'"))
            {
                result = result.Substring(0, result.Length - 1);
            }

            if (result.Length > 31)
            {
                result = result.Substring(0, 31);
            }

            result = result.Replace("\"", string.Empty);
            result = result.Replace("/", string.Empty);
            result = result.Replace("?", string.Empty);
            result = result.Replace("*", string.Empty);
            result = result.Replace("[", string.Empty);
            result = result.Replace("]", string.Empty);
            result = result.Replace(":", string.Empty);

            result = result.Replace("`", string.Empty);
            result = result.Replace("~", string.Empty);
            result = result.Replace("!", string.Empty);
            result = result.Replace("@", string.Empty);
            result = result.Replace("#", string.Empty);
            result = result.Replace("$", string.Empty);
            result = result.Replace("%", string.Empty);
            result = result.Replace("^", string.Empty);
            result = result.Replace("&", string.Empty);
            result = result.Replace("(", string.Empty);
            result = result.Replace(")", string.Empty);
            result = result.Replace("-", string.Empty);
            result = result.Replace("_", string.Empty);
            result = result.Replace("=", string.Empty);
            result = result.Replace("+", string.Empty);
            result = result.Replace("{", string.Empty);
            result = result.Replace("}", string.Empty);
            result = result.Replace("|", string.Empty);
            result = result.Replace(";", string.Empty);
            result = result.Replace(":", string.Empty);
            result = result.Replace(",", string.Empty);
            result = result.Replace("<", string.Empty);
            result = result.Replace(".", string.Empty);
            result = result.Replace(">", string.Empty);

            //result = Regex.Replace(result, @"[^0-9a-zA-Z]+", "");

            List<string> worksheetNames = workbook.WorksheetNames();
            if(worksheetNames != null && worksheetNames.Count != 0)
            {
                int index = startIndex;
                string name_Temp = result;
                while(worksheetNames.Contains(name_Temp))
                {
                    name_Temp = result + separator + index;
                    int length = 1;
                    while(name_Temp.Length > 31)
                    {
                        if(length >= result.Length)
                        {
                            return null;
                        }

                        name_Temp = result.Substring(0, result.Length - length) + separator + index;
                        length++;
                    }

                    index++;
                }

                result = name_Temp;
            }

            return result;
        }
    }
}