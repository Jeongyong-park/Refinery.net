using NPOI.POIFS.Crypt.Dsig;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Refinery.Result
{
    public class GenericParsedRecord : ParsedRecord
    {
        public GenericParsedRecord() { }

        public GenericParsedRecord(Dictionary<string, object> data)
        {
            this.ExtractedRawData = data;
        }


    }
}