using System;
using NPOI.SS.UserModel;

namespace Refinery.Cell.Parser
{
    public class IntCellParser : CellParser<int?>
    {
        public override int? TryParse(ICell cell)
        {
            if (cell == null) return null;

            int? value = null;

            if (GetConcreteCellType(cell) == CellType.Numeric && !DateUtil.IsCellInternalDateFormatted(cell))
            {
                var doubleValue = cell.NumericCellValue;
                if (doubleValue.Equals(Math.Round(doubleValue)))
                    value = (int)doubleValue;
            }
            else if (int.TryParse(cell.ToString().Trim(), out var result))
            {
                value = result;
            }

            return value;

        }

    }
}
