using System;
using System.Collections.Generic;
using System.Linq;

namespace Refinery.Exceptions
{
    public class ExceptionManager
    {
        private List<ExceptionData> exceptions = new List<ExceptionData>();

        public void Register(Exception exception, Location location = null)
        {
            switch (exception)
            {
                case ManagedException managedException:
                    exceptions.Add(new ExceptionData(managedException, location));
                    break;
                default:
                    exceptions.Add(new ExceptionData(new UncategorizedException(exception.ToString()), location));
                    break;
            }
        }

        public bool ContainsCritical()
        {
            return exceptions.Any(exceptionData => exceptionData.IsCritical());
        }

        public bool IsEmpty()
        {
            return exceptions.Count == 0;
        }

        public List<Dictionary<string, object>> ExtractData()
        {
            var sortedByCriticality = exceptions.OrderBy(exceptionData => exceptionData.Exception.Level);
            return sortedByCriticality.Select(exceptionData =>
            {
                var data = exceptionData.Exception.ExtractData();
                if (exceptionData.Location != null)
                {
                    data = MergeDictionaries(data, exceptionData.Location.ExtractData());
                }
                return data;
            }).ToList();
        }

        public List<ExceptionData> Exceptions
        {
            get { return this.exceptions; }
        }

        private Dictionary<string, object> MergeDictionaries(Dictionary<string, object> dictionary1, Dictionary<string, object> dictionary2)
        {
            var mergedDictionary = new Dictionary<string, object>(dictionary1);
            foreach (var keyValuePair in dictionary2)
            {
                mergedDictionary[keyValuePair.Key] = keyValuePair.Value;
            }
            return mergedDictionary;
        }

        public class Location
        {
            public string SheetName { get; }
            public int? RowNumber { get; }

            public Location(string sheetName, int? rowNumber = null)
            {
                SheetName = sheetName;
                RowNumber = rowNumber;
            }

            public Dictionary<string, object> ExtractData()
            {
                var data = new Dictionary<string, object>
            {
                { "sheetName", SheetName }
            };
                if (RowNumber != null)
                {
                    data.Add("rowNumber", RowNumber);
                }
                return data;
            }

            public override string ToString()
            {
                return $"SheetName: {SheetName}, rowNumber: {RowNumber}";
            }
        }

        public class ExceptionData
        {
            public ManagedException Exception { get; }
            public Location Location { get; }

            public ExceptionData(ManagedException exception, Location location = null)
            {
                Exception = exception;
                Location = location;
            }

            public bool IsCritical()
            {
                return Exception.Level == Level.CRITICAL;
            }
        }
    }
}