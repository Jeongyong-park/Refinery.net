using System.Collections.Generic;

namespace Refinery.Domain
{
    public class WorkbookParserDefinition
    {
        public List<SheetParserDefinition> SpreadsheetParserDefinitions { get; set; } = new List<SheetParserDefinition>() { };
        public bool IncludeHidden { get; set; } = false;

        public WorkbookParserDefinition() { }
        public WorkbookParserDefinition(List<SheetParserDefinition> spreadsheetParserDefinitions, bool includeHidden = false)
        {
            SpreadsheetParserDefinitions = spreadsheetParserDefinitions;
            IncludeHidden = includeHidden;
        }
    }
}