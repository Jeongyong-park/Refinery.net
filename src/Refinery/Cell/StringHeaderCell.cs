using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using Org.BouncyCastle.Crypto.Agreement.JPake;
using ICell = NPOI.SS.UserModel.ICell;

namespace Refinery.Cell
{
    /// <summary>
    /// Matches headers using case-insensitive substring comparison.
    /// </summary>
    public class StringHeaderCell : AbstractHeaderCell
    {
        /// <summary>
        /// The patterns to match.
        /// </summary>
        private List<string> Patterns { get; }

        public StringHeaderCell(List<string> patterns) { Patterns = patterns; }

        public StringHeaderCell(string pattern) : this(new List<string> { pattern }) { }

        public override bool Matches(ICell cell)
        {
            if (cell.CellType != CellType.String)
                return false;

            string cellValue = cell.StringCellValue.Trim();

            return Patterns.Any(pattern => cellValue.ToLower().Contains(pattern.ToLower()));

        }

        public override string ToString() => $"[{String.Join(", ", Patterns)}]";


    }
}
