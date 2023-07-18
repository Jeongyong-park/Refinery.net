using System;
using NPOI.SS.UserModel;
using System.Globalization;


namespace Refinery.Cell.Parser
{
    public class DateTimeFormatCellParser : CellParser<DateTime?>
    {
        private readonly string format;

        public DateTimeFormatCellParser(string format)
        {
            this.format = format;
        }

        public override DateTime? TryParse(ICell? cell)
        {
            if (cell == null) return null;

            DateTime? value;
            if (GetConcreteCellType(cell) == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            {
                value = cell.DateCellValue;
            }
            else
            {
                value = ParseStringWithFormatter(GetStringValue(cell), format);
            }

            return value;
        }

        private string GetStringValue(ICell cell)
        {
            return GetConcreteCellType(cell) == CellType.String ? cell.StringCellValue.Trim() : cell.ToString().Trim();
        }

        private DateTime? ParseStringWithFormatter(string dateStr, string format)
        {
            DateTime? result = null;
            try
            {
                result = DateTime.ParseExact(dateStr, format, null);
            }
            catch (FormatException)
            {
                // Ignore and return null
                result = null;

            }

            return result;
        }
    }

}