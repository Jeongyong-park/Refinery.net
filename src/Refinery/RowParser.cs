using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Refinery.Cell.Parser;
using Refinery.Cell;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery
{
    public abstract class RowParser
    {
        private RowParserData rowParserData;
        private ExceptionManager exceptionManager;
        public RowParser(RowParserData rowParserData, ExceptionManager exceptionManager)
        {
            this.rowParserData = rowParserData;
            this.exceptionManager = exceptionManager;
        }

        public abstract ParsedRecord ToRecord(IRow row);

        public ParsedRecord ToRecordOrDefault(IRow row)
        {
            try
            {
                return ToRecord(row);
            }
            catch (Exception ex)
            {
                // Handle the exception and return a default record
                // You can register the exception with an exception manager if needed
                // For simplicity, let's just return an empty GenericParsedRecord
                return new GenericParsedRecord(new Dictionary<string, object>());
            }
        }

        public bool Skip(IRow row)
        {
            // Optionally implement the logic to determine if the row should be skipped
            return false;
        }

        public ParsedRecord ExtractDataFromPreviousRecord(ParsedRecord current, ParsedRecord previous)
        {
            // Optionally implement the logic to extract data from the previous record
            return current;
        }

        public bool ShouldGroupRows(ParsedRecord current, ParsedRecord previous)
        {
            // Optionally implement the logic to determine if the rows should be grouped
            return false;
        }

        public virtual bool ShouldStoreExtractedRawDataInParsedRecord()
        {
            return true;
        }

        public Dictionary<string, object> ExtractAllData(IRow row)
        {
            var rowData = ExtractDataFromRow(row);
            Dictionary<string, object> metadata = rowParserData.Metadata.AllData();
            metadata.Add(Metadata.ROW_NUMBER, row.RowNum + 1);
            return MergeDictionaries(metadata, rowData);
        }

        public string ParseRequiredFieldAsString(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return new StringCellParser().Parse(cell);
        }

        public string ParseOptionalFieldAsString(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return new StringCellParser().TryParse(cell);
        }

        public double ParseRequiredFieldAsDouble(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return (double)new DoubleCellParser().Parse(cell);
        }

        public double? ParseOptionalFieldAsDouble(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return new DoubleCellParser().TryParse(cell);
        }

        public int ParseRequiredFieldAsInteger(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return (int)new IntCellParser().Parse(cell);
        }

        public int? ParseOptionalFieldAsInteger(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return new IntCellParser().TryParse(cell);
        }

        protected DateTime ParseRequiredFieldAsDateTime(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return (DateTime)new DateTimeCellParser().Parse(cell);
        }

        protected DateTime? ParseOptionalFieldAsDateTime(IRow row, AbstractHeaderCell headerCell)
        {
            ICell cell = FindCell(row, headerCell);
            return new DateTimeCellParser().TryParse(cell);
        }

        protected DateTime ParseRequiredFieldAsDateTime(IRow row, AbstractHeaderCell headerCell, string format)
        {
            ICell cell = FindCell(row, headerCell);
            return (DateTime)new DateTimeFormatCellParser(format).Parse(cell);
        }

        protected DateTime? ParseOptionalFieldAsDateTime(IRow row, AbstractHeaderCell headerCell, string format)
        {
            ICell cell = FindCell(row, headerCell);
            return new DateTimeFormatCellParser(format).TryParse(cell);
        }

        private ICell FindCell(IRow row, AbstractHeaderCell headerCell)
        {
            this.rowParserData.HeaderMap.TryGetValue(headerCell, out int cellIndex);
            return FindCellByIndex(row, cellIndex);
        }

        private ICell FindCellByIndex(IRow row, int cellIndex)
        {
            var maybeMergedCell = this.rowParserData.MergedCellsResolver[row.RowNum, cellIndex];

            return maybeMergedCell ?? row.GetCell(cellIndex, MissingCellPolicy.RETURN_BLANK_AS_NULL);

        }

        private Dictionary<string, object> ExtractDataFromRow(IRow row)
        {
            Dictionary<string, object> rowData = new Dictionary<string, object>();
            foreach (KeyValuePair<string, int> headerCell in rowParserData.AllHeadersMapping)
            {
                KeyValuePair<string, object>? cellData = ResolveCellValue(headerCell, row);
                if (cellData != null)
                {
                    rowData.Add(cellData.Value.Key, cellData.Value.Value);
                }
            }
            return rowData;
        }

        private KeyValuePair<string, object>? ResolveCellValue(KeyValuePair<string, int> headerCell, IRow row)
        {
            ICell cell = FindCellByIndex(row, headerCell.Value);
            if (cell == null)
            {
                return null;
            }

            object value = GetCellValue(cell);
            if (value == null)
            {
                return null;
            }

            return new KeyValuePair<string, object>(headerCell.Key, value);
        }
        private object GetCellValue(ICell cell)
        {
            CellType cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
            switch (cellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    else
                    {
                        double doubleValue = cell.NumericCellValue;
                        return doubleValue == Math.Round(doubleValue) ? (int)doubleValue : doubleValue;
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                default:
                    return null;
            }
        }

        private static CellRangeAddress GetMergedRegion(ISheet sheet, int rowNum, int colNum)
        {
            int numMergedRegions = sheet.NumMergedRegions;
            for (int i = 0; i < numMergedRegions; i++)
            {
                CellRangeAddress mergedRegion = sheet.GetMergedRegion(i);
                if (mergedRegion.IsInRange(rowNum, colNum))
                {
                    return mergedRegion;
                }
            }
            return null;
        }

        private Dictionary<string, object> MergeDictionaries(Dictionary<string, object> dictionary1, Dictionary<string, object> dictionary2)
        {
            Dictionary<string, object> mergedDictionary = new Dictionary<string, object>(dictionary1);
            foreach (KeyValuePair<string, object> kvp in dictionary2)
            {
                if (!mergedDictionary.ContainsKey(kvp.Key))
                {
                    mergedDictionary.Add(kvp.Key, kvp.Value);
                }
            }
            return mergedDictionary;
        }
    }
}
