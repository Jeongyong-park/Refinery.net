using Refinery.Cell.Parser;

namespace Refinery.Tests.Cell.Parser
{
    public class DateTimeCellParserTest : CellParserTest
    {
        private readonly CellParser<DateTime?> _parser = new DateTimeCellParser();

        [Fact]
        public void ShouldParseToDateTime()
        {
            // expect
            Assert.Equal(Convert.ToDateTime("2021-12-12T00:00"), _parser.Parse(DateCell()));
        }

        [Fact]
        public void ShouldTryToParseToDateTimeOrNullIfFailed()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCell()));
            Assert.Null(_parser.TryParse(EmptyStringCell()));
            Assert.Null(_parser.TryParse(DoubleCell()));
            Assert.Null(_parser.TryParse(IntCell()));
            Assert.Equal(Convert.ToDateTime("2021-12-12T00:00"), _parser.TryParse(DateCell()));
            Assert.Null(_parser.TryParse(BoolCell()));
            Assert.Null(_parser.TryParse(NullCell()));
            Assert.Null(_parser.TryParse(DoubleAsStringCell()));
            Assert.Null(_parser.TryParse(DoubleInt()));
        }

        [Fact]
        public void ShouldThrowExceptionIfFailedToParseToDateTime()
        {
            // expect
            Assert.Throws<CellParserException>(() => _parser.Parse(StringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(EmptyStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(BoolCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(NullCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DoubleCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DoubleAsStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(DoubleInt()));
            Assert.Throws<CellParserException>(() => _parser.Parse(IntCell()));
        }

        [Fact]
        public void ShouldParseToDateTimeFromFormula()
        {
            // expect
            Assert.Equal(Convert.ToDateTime("2021-12-12T00:00"), _parser.Parse(DateCellFromFormula()));
        }

        [Fact]
        public void ShouldTryToParseToDateTimeOrNullIfFailedFromFormula()
        {
            // expect
            Assert.Null(_parser.TryParse(StringCellFromFormula()));
            Assert.Null(_parser.TryParse(EmptyStringCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleCellFromFormula()));
            Assert.Null(_parser.TryParse(IntCellFromFormula()));
            Assert.Equal(Convert.ToDateTime("2021-12-12T00:00"), _parser.TryParse(DateCellFromFormula()));
            Assert.Null(_parser.TryParse(BoolCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleAsStringCellFromFormula()));
            Assert.Null(_parser.TryParse(DoubleIntFromFormula()));
            Assert.Equal(Convert.ToDateTime("2021-12-12T15:43:43"), _parser.TryParse(DateTimeFromFormula()));
        }
    }
}
