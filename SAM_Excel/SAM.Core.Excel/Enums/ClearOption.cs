using System.ComponentModel;

namespace SAM.Core.Excel
{
    [Description("Content Clear Option")]
    public enum ClearOption
    {
        [Description("Undefined")] Undefined,
        [Description("None")] None,
        [Description("All")] All,
        [Description("Data")] Data,
        [Description("Column")] Column,
    }
}