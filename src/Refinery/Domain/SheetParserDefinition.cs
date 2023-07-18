using System;
using System.Collections.Generic;

namespace Refinery.Domain
{
    public class SheetParserDefinition
    {
        public Func<string, bool> SheetNameFilter { get; }
        public List<TableParserDefinition> TableDefinitions { get; }
        public List<MetadataEntryDefinition> MetadataParserDefinition { get; }

        public SheetParserDefinition(Func<string, bool> sheetNameFilter, List<TableParserDefinition> tableDefinitions, List<MetadataEntryDefinition> metadataParserDefinition = null)
        {
            SheetNameFilter = sheetNameFilter;
            TableDefinitions = tableDefinitions;
            MetadataParserDefinition = metadataParserDefinition ?? new List<MetadataEntryDefinition>();
        }
    }
}
