using NPOI.SS.UserModel;
using System;
using System.ComponentModel.Design;

namespace Refinery.Cell.Parser
{
    public class DateTimeCellParser : CellParser<DateTime?>
    {

        public override DateTime? TryParse(ICell? cell)
        {
            if (cell == null) return null;//default(DateTime);
            
            DateTime? value = null;
            if (GetConcreteCellType(cell) == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                value = cell.DateCellValue;
     
            return value;



        }
    }
}
