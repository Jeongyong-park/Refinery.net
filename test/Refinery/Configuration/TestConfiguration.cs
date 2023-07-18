using Refinery.Cell;
using Refinery.Domain;

namespace Refinery.Tests.Configuration
{
    public static class TestConfiguration
    {
        public static StringHeaderCell String = new StringHeaderCell("string");
        public static StringHeaderCell Number = new StringHeaderCell("number");
        public static StringHeaderCell Date = new StringHeaderCell("date");
        public static StringHeaderCell OptionalString = new StringHeaderCell("optional_str");

        public static WorkbookParserDefinition TestDefinition = new WorkbookParserDefinition
        {
            SpreadsheetParserDefinitions = new List<SheetParserDefinition>
        {
            new SheetParserDefinition
            (
                (sheetName) => true,
                new List<TableParserDefinition>
                {
                    new TableParserDefinition(
                        new HashSet<AbstractHeaderCell>
                        {
                            String,
                            Number,
                            Date,
                            OptionalString
                        },
                        new HashSet<AbstractHeaderCell>{ }

                        )

                }
            )
        }
        };
    }
}