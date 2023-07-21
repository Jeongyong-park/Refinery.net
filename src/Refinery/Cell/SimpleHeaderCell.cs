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

        public SimpleHeaderCell(string name)
        {
            this.Name = name;
        }

        public override bool Matches(ICell cell) => cell.ToString().Trim().Equals(Name) ;


    }
}
