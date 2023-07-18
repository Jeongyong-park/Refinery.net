using NPOI.SS.UserModel;
using System;

namespace Refinery.Cell.Parser
{
    public abstract class CellParser<T>
    {
        public T Parse(ICell cell = null)
        {
            T value = TryParse(cell);
            if (value == null)
            {
                throw new CellParserException("Cell is either empty or not parsable");
            }
            return value;
        }

        public abstract T TryParse(ICell cell = null);

        protected CellType GetConcreteCellType(ICell cell)
        {
            return cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        }
    }

    public class CellParserException : Exception
    {
        public CellParserException(string message) : base(message)
        {
        }
    }
}
