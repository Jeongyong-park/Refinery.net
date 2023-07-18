using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace Refinery.Cell
{
    /// <summary>
    /// Matches headers using exact string comparison.
    /// </summary>
    public class SimpleHeaderCell : AbstractHeaderCell
    {
        /// <summary>
        /// The header name.
        /// </summary>
        public string Name { get; }

        public SimpleHeaderCell(string name)
        {
            Name = name;
        }

        public override bool Matches(ICell cell) => cell.ToString().Trim().Equals(Name) ;

        public override string ToString() => Name;
    }
}
