using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Refinery.Cell
{
    public abstract class AbstractHeaderCell
    {
        public abstract bool Matches(ICell cell);

        public bool Inside(IEnumerable<ICell> values)
        {
            return values.Any(cell => Matches(cell));
        }
    }
}
