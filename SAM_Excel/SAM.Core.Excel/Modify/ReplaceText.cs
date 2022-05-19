using NetOffice.ExcelApi;
using System.Collections.Generic;
using System.Linq;

namespace SAM.Core.Excel
{
    public static partial class Modify
    {
        public static List<string> ReplaceText<T>(this Range range, Dictionary<string, T> dictionary)
        {
            if(range == null || dictionary == null)
            {
                return null;
            }

            object[,] values = range.Value as object[,];
            if(values == null)
            {
                return null;
            }

            HashSet<string> hashSets = new HashSet<string>();
            if(dictionary.Count != 0)
            {
                for (int i = values.GetLowerBound(0); i < values.GetUpperBound(0) + 1; i++)
                {
                    for (int j = values.GetLowerBound(1); j < values.GetUpperBound(1) + 1; j++)
                    {
                        if (values[i, j] == null)
                        {
                            continue;
                        }

                        string text = values[i, j]?.ToString();
                        if(text == null)
                        {
                            continue;
                        }

                        if(!dictionary.TryGetValue(text, out T value))
                        {
                            continue;
                        }

                        values[i, j] = value;
                        hashSets.Add(text);
                    }
                }

                if (hashSets.Count != 0)
                {
                    range.Value = values;
                }
            }

            return hashSets.ToList();
        }
    }
}


