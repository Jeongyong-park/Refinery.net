using Refinery.Cell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Refinery.Tests.Cell
{
    public class TestMergedCellsResolver
    {
        [Fact]
        public void ShouldThrowIfCellRowIndexIsNegative()
        {
            // expect
            Assert.Throws<ArgumentException>(() => new MergedCellsResolver.CellLocation(-1, 0));
        }

        [Fact]
        public void ShouldThrowIfCellColumnIndexIsNegative()
        {
            // expect
            Assert.Throws<ArgumentException>(() => new MergedCellsResolver.CellLocation(0, -1));
        }
    }
}
