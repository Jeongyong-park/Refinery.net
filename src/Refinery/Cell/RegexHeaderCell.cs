using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Refinery.Cell
{
    /// <summary>
    /// Matches headers using a regex or a list of regexes.
    /// </summary>
    public class RegexHeaderCell : AbstractHeaderCell
    {
        /// <summary>
        /// The regex patterns.
        /// </summary>
        public List<Regex> Patterns { get; }

        public RegexHeaderCell(List<Regex> patterns)
        {
            Patterns = patterns;
        }

        public RegexHeaderCell(Regex pattern) : this(new List<Regex> { pattern })
        {
        }

        public RegexHeaderCell(string pattern) : this(new List<Regex> { new Regex(pattern) })
        {
        }

        public override bool Matches(ICell cell)
        {
            if (cell.CellType != CellType.String)
                return false;

            return Patterns.Any(pattern => pattern.IsMatch(cell.StringCellValue));
        }

        public override string ToString() => Patterns.ToString();
    }
}
