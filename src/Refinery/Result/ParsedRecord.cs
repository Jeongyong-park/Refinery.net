using System;
using System.Collections.Generic;
using System.Linq;


namespace Refinery.Result
{

    [Serializable]
    public abstract class ParsedRecord
    {
        public Guid? GroupId { get; set; }

        [field: NonSerialized]
        public Dictionary<string, object> ExtractedRawData { get; set; }

        public void CloneRawData(ParsedRecord from)
        {
            ExtractedRawData = from.ExtractedRawData;
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(this, obj)) return true;
            if (GetType() != obj?.GetType()) return false;

            ParsedRecord other = (ParsedRecord)obj;

            var a = ExtractedRawData
                .Where(d => d.Key != "anchor")
                // .OrderBy(e => e.Key)
                .ToDictionary(
                    item => item.Key,
                    item => GetValueForComparison(item.Value));

            var b = other.ExtractedRawData
                .Where(d => d.Key != "anchor")
                // .OrderBy(e => e.Key)
                .ToDictionary(
                    item => item.Key,
                    item => GetValueForComparison(item.Value));

            return GroupId == other.GroupId && !a.Except(b).Any();
        }

        private object GetValueForComparison(object value)
        {
            if (value is int || value is float || value is short || value is long)
                return Convert.ToDouble(value);

            return value;
        }

        public override int GetHashCode()
        {
            int hashCode = GroupId?.GetHashCode() ?? 0;
            hashCode = (hashCode * 31) + (ExtractedRawData?.GetHashCode() ?? 0);
            return hashCode;
        }

        public override string ToString()
        {
            return $"ParsedRecord(groupId={GroupId}, extractedRawData={ExtractedRawData})";
        }

    }
}