using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace Refinery.Cell
{
    /// <summary>
    /// Can be used to define the order of a header cell.
    /// Use this to differentiate if you have the same header mentioned twice.
    /// `new OrderedHeaderCell(new SimpleHeaderCell("header_pattern"), 1)`
    /// 
    /// You can also have `MergedHeaderCell` as a target.
    /// </summary>
    public class OrderedHeaderCell : AbstractHeaderCell
    {
        /// <summary>
        /// The header cell.
        /// </summary>
        public AbstractHeaderCell HeaderCell { get; }

        /// <summary>
        /// The priority. Lower value means it comes first.
        /// </summary>
        public int Priority { get; }

        public OrderedHeaderCell(AbstractHeaderCell headerCell, int priority)
        {
            HeaderCell = headerCell;
            Priority = priority;
        }

        public override bool Matches(ICell cell) => HeaderCell.Matches(cell);

        public override string ToString() => HeaderCell.ToString();
    }
}
