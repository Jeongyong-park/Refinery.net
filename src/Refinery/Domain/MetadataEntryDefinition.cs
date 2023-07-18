using NPOI.SS.UserModel;

namespace Refinery.Domain
{
    public class MetadataEntryDefinition
    {
        public string MetadataName { get; }
        public string MatchingCellKey { get; }
        public MetadataValueLocation ValueLocation { get; }
        public System.Func<ICell, object> Extractor { get; }

        public MetadataEntryDefinition(string metadataName, string matchingCellKey, MetadataValueLocation valueLocation, System.Func<ICell, object> extractor)
        {
            MetadataName = metadataName;
            MatchingCellKey = matchingCellKey;
            ValueLocation = valueLocation;
            Extractor = extractor;
        }
    }
}
