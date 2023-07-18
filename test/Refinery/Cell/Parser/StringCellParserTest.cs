using Refinery.Cell.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Refinery.Tests.Cell.Parser
{
    public class StringCellParserTest : CellParserTest
    {
        private readonly CellParser<string> _parser = new StringCellParser();

        [Fact]
        public void ShouldParseToString()
        {
            // expect
            Assert.Equal("example", _parser.Parse(StringCell()));
            Assert.Equal("3.1415", _parser.Parse(DoubleCell()));
            Assert.Equal("1", _parser.Parse(IntCell()));
            //Assert.Equal("12-Dec-2021", parser.TryParse(DateCell()));
            Assert.Matches(@"^12-\w{3}-2021$", _parser.Parse(DateCell()));
            Assert.Equal("TRUE", _parser.Parse(BoolCell()));
            Assert.Equal("2.78", _parser.Parse(DoubleAsStringCell()));
        }

        [Fact]
        public void ShouldTryToParseToStringOrNullIfFailed()
        {
            // expect
            Assert.Equal("example", _parser.TryParse(StringCell()));
            Assert.Null(_parser.TryParse(EmptyStringCell()));
            Assert.Equal("3.1415", _parser.TryParse(DoubleCell()));
            Assert.Equal("1", _parser.TryParse(IntCell()));
            //Assert.Equal("12-Dec-2021", parser.TryParse(DateCell()));
            Assert.Matches(@"^12-\w{3}-2021$", _parser.Parse(DateCell()));
            Assert.Equal("TRUE", _parser.TryParse(BoolCell()));
            Assert.Null(_parser.TryParse(NullCell()));
            Assert.Equal("2.78", _parser.TryParse(DoubleAsStringCell()));
        }

        [Fact]
        public void ShouldThrowExceptionIfFailedToParseToString()
        {
            // expect
            Assert.Throws<CellParserException>(() => _parser.Parse(EmptyStringCell()));
            Assert.Throws<CellParserException>(() => _parser.Parse(NullCell()));
        }

        [Fact]
        public void ShouldParseToStringFromFormula()
        {
            // expect
            Assert.Equal("example", _parser.Parse(StringCellFromFormula()));
            Assert.Equal("2.78", _parser.Parse(DoubleAsStringCellFromFormula()));
            Assert.Equal("15-11-2021", _parser.Parse(DateStrFromFormula()));
            Assert.Equal("15-11-2021 @ 15:43", _parser.Parse(DateTimeStrFromFormula()));
        }

        [Fact]
        public void ShouldTryToParseToStringOrNullIfFailedFromFormula()
        {
            // expect
            Assert.Equal("example", _parser.TryParse(StringCellFromFormula()));
            Assert.Equal("2.78", _parser.TryParse(DoubleAsStringCellFromFormula()));
            Assert.Equal("15-11-2021", _parser.TryParse(DateStrFromFormula()));
            Assert.Equal("15-11-2021 @ 15:43", _parser.TryParse(DateTimeStrFromFormula()));
            Assert.Null(_parser.TryParse(EmptyStringCellFromFormula()));
        }
    }
}
