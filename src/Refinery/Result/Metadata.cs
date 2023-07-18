using System.Collections.Generic;

namespace Refinery.Result
{
    public class Metadata
    {
        private readonly Dictionary<string, object> data;
        private string divider = string.Empty;

        public Metadata(Dictionary<string, object> data)
        {
            this.data = data;
        }

        public string GetDivider()
        {
            return divider;
        }

        public void SetDivider(string value)
        {
            this.divider = value;
        }

        public string GetWorkbookName()
        {
            return (string)data[WORKBOOK_NAME];
        }

        public string GetSheetName()
        {
            return (string)data[SPREADSHEET_NAME];
        }

        public string GetAnchor()
        {
            return (string)data[ANCHOR];
        }

        public Metadata Plus(KeyValuePair<string, object> element)
        {
            var newData = new Dictionary<string, object>(data);
            newData[element.Key] = element.Value;
            return new Metadata(newData);
        }

        public object this[string key]
        {
            get { return data[key]; }
        }

        public Dictionary<string, object> AllData()
        {
            if (divider == null)
            {
                return data;
            }
            else
            {
                var newData = new Dictionary<string, object>(data);
                newData[DIVIDER] = divider;
                return newData;
            }
        }

        public const string WORKBOOK_NAME = "workbook_name";
        public const string SPREADSHEET_NAME = "spreadsheet_name";
        public const string ANCHOR = "anchor";
        public const string DIVIDER = "divider";
        public const string ROW_NUMBER = "row_number";
    }
}
