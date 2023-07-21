using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using NPOI.SS.UserModel;

namespace Refinery.Cell
{
    public abstract class AbstractHeaderCell
    {
        /// <summary>
        /// The header name.
        /// </summary>
        public string Name { get; protected set; }

        public abstract bool Matches(ICell cell);

        public bool Inside(IEnumerable<ICell> values)
        {
            return values.Any(cell => Matches(cell));
        }

        public override string ToString() => Name;
    }
}
