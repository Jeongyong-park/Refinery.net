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
        public string defaultName { get; protected set; }

        public StringHeaderCell(List<string> patterns, string defaultName = "") { Patterns = patterns; this.defaultName = defaultName; }

        public StringHeaderCell(string pattern, string defaultName = "") : this(new List<string> { pattern }, defaultName) { }
        public StringHeaderCell(string pattern) : this(pattern, pattern) { }

        public override bool Matches(ICell cell)
        {
            if (cell.CellType != CellType.String)
                return false;

            string cellValue = cell.StringCellValue.Trim();

            return Patterns.Any(pattern => cellValue.ToLower().Contains(pattern.ToLower()));

        }

        public override string ToString() =>
            !string.IsNullOrWhiteSpace(defaultName) ? defaultName : $"[{String.Join(", ", Patterns)}]";


    }
}
