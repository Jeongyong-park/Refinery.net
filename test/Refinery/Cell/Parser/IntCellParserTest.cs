using Refinery.Cell.Parser;
using Xunit;

namespace Refinery.Tests.Cell.Parser
{
    public class IntCellParserTest : CellParserTest
    {
        private CellParser<int?> _parser = new IntCellParser();

        [Fact]
        public void ShouldParseToInt()
        {
            // expect
            Assert.Equal(1, _parser.Parse(IntCell()));
            Assert.Equal(3, _parser.Parse(DoubleInt()));
        }

        [Fact]
        public void ShouldTryParseToIntOrNullIfFailed()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCell()));
            Assert.Null(_parser.TryParse(EmptyStringCell()));
            Assert.Null(_parser.TryParse(DoubleCell()));
            Assert.Equal(1, _parser.TryParse(IntCell()));
            Assert.Null(_parser.TryParse(DateCell()));
            Assert.Null(_parser.TryParse(BoolCell()));
            Assert.Null(_parser.TryParse(NullCell()));
            Assert.Null(_parser.TryParse(DoubleAsStringCell()));
            Assert.Equal(3, _parser.TryParse(DoubleInt()));
        }

        [Fact]
        public void ShouldThrowExceptionIfFailedToParseToInt()
        {
            // expect
            Assert.Throws<CellParserException>(() => _parser.Parse(StringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(EmptyStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DoubleCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DoubleAsStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DateCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(BoolCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(NullCell()));
        }

        [Fact]
        public void ShouldParseToIntFromFormula()
        {
            // expect
            Assert.Equal(1, _parser.Parse(IntCellFromFormula()));
            Assert.Equal(3, _parser.Parse(DoubleIntFromFormula()));
        }

        [Fact]
        public void ShouldTryParseToIntOrNullIfFailedFromFormula()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCellFromFormula()));
            Assert.Null(_parser.TryParse(EmptyStringCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleCellFromFormula()));
            Assert.Equal(1, _parser.TryParse(IntCellFromFormula()));
            Assert.Null(_parser.TryParse(DateCellFromFormula()));
            Assert.Null(_parser.TryParse(BoolCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleAsStringCellFromFormula()));
            Assert.Equal(3, _parser.TryParse(DoubleIntFromFormula()));
        }
    }
}
