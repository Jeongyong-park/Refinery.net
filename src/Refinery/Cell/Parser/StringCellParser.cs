using System;
using NPOI.SS.UserModel;

namespace Refinery.Cell.Parser
{
    public class StringCellParser : CellParser<string>
    {
        public override string TryParse(ICell cell)
        {
            if (cell == null) return null;

            string val = GetConcreteCellType(cell).Equals(CellType.String) ? cell.StringCellValue.Trim() : cell.ToString().Trim();

            return !String.IsNullOrWhiteSpace(val) ? val : null;
        }
    }
}
