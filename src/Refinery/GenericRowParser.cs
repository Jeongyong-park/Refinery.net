using System.Collections.Generic;
using NPOI.SS.UserModel;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery
{
    public class GenericRowParser : RowParser
    {
        public GenericRowParser(RowParserData rowParserData, ExceptionManager exceptionManager)
            : base(rowParserData, exceptionManager)
        {
        }

        public override ParsedRecord ToRecord(IRow row)
        {
            Dictionary<string, object> allData = ExtractAllData(row);
            return new GenericParsedRecord(allData);
        }

        public override bool ShouldStoreExtractedRawDataInParsedRecord()
        {
            return false;
        }
    }
}