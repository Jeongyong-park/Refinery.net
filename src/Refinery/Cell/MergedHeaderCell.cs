using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace Refinery.Cell
{
    /// <summary>
    /// Can be used to parse a merged header cell
    /// </summary>
    /// <seealso cref="com.vortexa.refinery.cell.OrderedHeaderCell"/>
    public class MergedHeaderCell : AbstractHeaderCell
    {
        /// <summary>
        /// Matching style. Any of <see cref="SimpleHeaderCell"/>, <see cref="StringHeaderCell"/>, or <see cref="RegexHeaderCell"/>
        /// </summary>
        public AbstractHeaderCell HeaderCell { get; }

        /// <summary>
        /// What cells to map to. Any of <see cref="SimpleHeaderCell"/>, <see cref="StringHeaderCell"/>, or <see cref="RegexHeaderCell"/>
        /// </summary>
        public HashSet<AbstractHeaderCell> HeaderCells { get; }

        public MergedHeaderCell(AbstractHeaderCell headerCell, HashSet<AbstractHeaderCell> headerCells)
        {
            HeaderCell = headerCell;
            HeaderCells = headerCells;

            if (headerCell is OrderedHeaderCell)
            {
                throw new ArgumentException("Use OrderedHeaderCell first instead");
            }
        }

        public override bool Matches(ICell cell) => HeaderCell.Matches(cell);

        public override string ToString() => HeaderCells.ToString();
    }
}
