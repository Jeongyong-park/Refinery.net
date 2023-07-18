using Refinery.Cell.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Refinery.Tests.Cell.Parser
{
    public class DoubleCellParserTest : CellParserTest
    {
        private readonly CellParser<double?> _parser = new DoubleCellParser();

        [Fact]
        public void ShouldParseToDouble()
        {
            // expect
            Assert.Equal(3.1415, _parser.Parse(DoubleCell()));
            Assert.Equal(1.0, _parser.Parse(IntCell()));
            Assert.Equal(2.78, _parser.Parse(DoubleAsStringCell()));
            Assert.Equal(3.0, _parser.Parse(DoubleInt()));
        }

        [Fact]
        public void ShouldTryToParseToDoubleOrNullIfFailed()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCell()));
            Assert.Null(_parser.TryParse(EmptyStringCell()));
            Assert.Equal(3.1415, _parser.TryParse(DoubleCell()));
            Assert.Equal(1.0, _parser.TryParse(IntCell()));
            Assert.Null(_parser.TryParse(DateCell()));
            Assert.Null(_parser.TryParse(BoolCell()));
            Assert.Null(_parser.TryParse(NullCell()));
            Assert.Equal(2.78, _parser.TryParse(DoubleAsStringCell()));
            Assert.Equal(3.0, _parser.TryParse(DoubleInt()));
        }

        [Fact]
        public void ShouldThrowExceptionIfFailedToParseToDouble()
        {
            // expect
            Assert.Throws<CellParserException>(() => _parser.Parse(StringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(EmptyStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(BoolCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(NullCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DateCell()));
        }

        [Fact]
        public void ShouldParseToDoubleFromFormula()
        {
            // expect
            Assert.Equal(3.1415, _parser.Parse(DoubleCellFromFormula()));
            Assert.Equal(1.0, _parser.Parse(IntCellFromFormula()));
            Assert.Equal(3.0, _parser.Parse(DoubleIntFromFormula()));
        }

        [Fact]
        public void ShouldTryToParseToDoubleOrNullIfFailedFromFormula()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCellFromFormula()));
            Assert.Null(_parser.TryParse(EmptyStringCellFromFormula()));
            Assert.Equal(3.1415, _parser.TryParse(DoubleCellFromFormula()));
            Assert.Equal(1.0, _parser.TryParse(IntCellFromFormula()));
            Assert.Null(_parser.TryParse(DateCellFromFormula()));
            Assert.Null(_parser.TryParse(BoolCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleAsStringCellFromFormula())); // Cannot parse double formula as string
            Assert.Equal(3.0, _parser.TryParse(DoubleIntFromFormula()));
        }
    }
}
