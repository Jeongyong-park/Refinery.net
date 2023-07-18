using System.Collections.Generic;
using Refinery.Cell;

namespace Refinery.Result
{
    public class RowParserData
    {
        public IDictionary<AbstractHeaderCell, int> HeaderMap { get; }
        public MergedCellsResolver MergedCellsResolver { get; }
        public Metadata Metadata { get; }
        public IDictionary<string, int> AllHeadersMapping { get; }

        public RowParserData(IDictionary<AbstractHeaderCell, int> headerMap, MergedCellsResolver mergedCellsResolver, Metadata metadata, IDictionary<string, int> allHeadersMapping)
        {
            HeaderMap = headerMap;
            MergedCellsResolver = mergedCellsResolver;
            Metadata = metadata;
            AllHeadersMapping = allHeadersMapping;
        }
    }
}
