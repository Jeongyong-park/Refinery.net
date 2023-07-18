using NPOI.SS.UserModel;

namespace Refinery.Cell.Parser
{
    public class DoubleCellParser : CellParser<double?>
    {
        public override double? TryParse(ICell cell)
        {
            if (cell == null) return null;

            double? value = null;

            if (GetConcreteCellType(cell) == CellType.Numeric && !DateUtil.IsCellInternalDateFormatted(cell))
            {
                value = cell.NumericCellValue;
            }
            else if (double.TryParse(cell.ToString().Trim(), out var result))
            {
                value = result;
            }

            return value;

        }
    }
}