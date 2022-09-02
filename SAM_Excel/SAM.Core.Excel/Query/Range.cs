using NetOffice.ExcelApi;

namespace SAM.Core.Excel
{
    public static partial class Query
    {
        public static Range Range(this Range range, string value)
        {
            if (range == null || string.IsNullOrEmpty(value))
            {
                return null;
            }

            object[,] values = range.Value as object[,];
            if (values == null || values.GetLength(0) == 0 || values.GetLength(1) == 0)
            {
                return null;
            }


            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
            {
                for (int j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++)
                {
                    object @object = values[i, j];
                    if (@object is string)
                    {
                        if ((string)@object == value)
                        {
                            Range range_Temp = range.Cells[i, j];
                            return range.Range(range_Temp, range_Temp.MergeArea);
                        }
                    }
                }
            }

            return null;
        }
    }
}