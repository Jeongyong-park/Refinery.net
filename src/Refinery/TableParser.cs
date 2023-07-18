using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using Refinery;
using Refinery.Cell;
using Refinery.Domain;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery
{
    internal class TableParser
    {
        private readonly ISheet sheet;
        private readonly TableParserDefinition definition;
        private readonly Metadata metadata;
        private readonly TableLocation location;
        private readonly MergedCellsResolver mergedCellsResolver;
        private readonly ExceptionManager exceptionManager;
        private readonly HeaderRowResolver headerRowResolver;

        public TableParser(
            ISheet sheet,
            TableParserDefinition definition,
            Metadata metadata,
            TableLocation location,
            MergedCellsResolver mergedCellsResolver,
            ExceptionManager exceptionManager,
            HeaderRowResolver headerRowResolver
        )
        {
            this.sheet = sheet;
            this.definition = definition;
            this.metadata = metadata;
            this.location = location;
            this.mergedCellsResolver = mergedCellsResolver;
            this.exceptionManager = exceptionManager;
            this.headerRowResolver = headerRowResolver;
        }

        public List<ParsedRecord> Parse()
        {
            IRow headerRow = FindHeaderRow();
            if (headerRow == null)
                return new List<ParsedRecord>();

            Dictionary<AbstractHeaderCell, int> columnHeaders = headerRowResolver.ResolveHeaderCellIndex(headerRow, definition.AllColumns());
            Dictionary<string, int> allHeadersMapping = MapAllHeaders(headerRow);
            TableLocationWithHeader tableLocationWithHeader = new TableLocationWithHeader(location.MinRow, headerRow.RowNum, location.MaxRow, allHeadersMapping.Values);
            CheckUncapturedHeaders(columnHeaders, allHeadersMapping, tableLocationWithHeader);
            Metadata enrichedMetadata = EnrichMetadata();

            if (definition.HasDivider)
            {
                return ParseTableWithDividers(columnHeaders, enrichedMetadata, tableLocationWithHeader, allHeadersMapping);
            }
            else
            {
                var rowParserData = new RowParserData(columnHeaders, mergedCellsResolver, enrichedMetadata, allHeadersMapping);
                RowParser rowParser = definition.RowParserFactory.Invoke(rowParserData, exceptionManager);
                return ParseTableWithoutDividers(rowParser, tableLocationWithHeader);
            }
        }

        private void CheckUncapturedHeaders(Dictionary<AbstractHeaderCell, int> columnHeaders, Dictionary<string, int> allHeadersMapping, TableLocationWithHeader tableLocationWithHeader)
        {
            HashSet<int> capturedColumnIndexes = new HashSet<int>(columnHeaders.Values);
            List<UncapturedHeadersException.UncapturedHeaderCell> uncapturedColumns = allHeadersMapping
                .Where(kvp => !capturedColumnIndexes.Contains(kvp.Value))
                .Select(kvp => new UncapturedHeadersException.UncapturedHeaderCell(kvp.Key, kvp.Value))
                .ToList();

            if (uncapturedColumns.Count > 0)
            {
                exceptionManager.Register(new UncapturedHeadersException(uncapturedColumns),
                    new ExceptionManager.Location(sheet.SheetName, tableLocationWithHeader.HeaderRow + 1));
            }
        }

        private List<ParsedRecord> ParseTableWithoutDividers(RowParser rowParser, TableLocationWithHeader location)
        {
            List<ParsedRecord> parsedRecords = new List<ParsedRecord>();

            foreach (int rowIndex in location.Range())
            {
                IRow row = sheet.GetRow(rowIndex);
                if (IsExtractableRow(row) && !rowParser.Skip(row))
                {
                    ParseRecordAndAddToResults(rowParser, row, parsedRecords);
                }
            }

            return parsedRecords;
        }

        private List<ParsedRecord> ParseTableWithDividers(Dictionary<AbstractHeaderCell, int> columnHeaders, Metadata enrichedMetadata, TableLocationWithHeader location, Dictionary<string, int> allHeadersMapping)
        {
            List<ParsedRecord> parsedRecords = new List<ParsedRecord>();
            RowParser rowParser = definition.RowParserFactory.Invoke(new RowParserData(columnHeaders, mergedCellsResolver, enrichedMetadata, allHeadersMapping), exceptionManager);

            foreach (int rowIndex in location.Range())
            {
                IRow row = sheet.GetRow(rowIndex);
                if (IsExtractableRow(row))
                {
                    if (IsDivider(row, location))
                    {
                        if (IsAllowedDivider(row, location, definition.AllowedDividers))
                        {
                            enrichedMetadata.SetDivider(FirstFilteredCell(row, location).ToString());
                        }
                    }
                    else if (!rowParser.Skip(row))
                    {
                        ParseRecordAndAddToResults(rowParser, row, parsedRecords);
                    }
                }
            }

            return parsedRecords;
        }

        private void ParseRecordAndAddToResults(RowParser rowParser, IRow row, List<ParsedRecord> parsedRecords)
        {
            ParsedRecord record = rowParser.ToRecordOrDefault(row);

            if (record is GenericParsedRecord)
            {
                parsedRecords.Add(record);
            }
            else
            {
                if (parsedRecords.Count == 0 || parsedRecords.Last() is GenericParsedRecord)
                {
                    parsedRecords.Add(record);
                }
                else
                {
                    ParsedRecord recordToAdd = rowParser.ExtractDataFromPreviousRecord(record, parsedRecords.Last());
                    recordToAdd.CloneRawData(record);

                    if (rowParser.ShouldGroupRows(recordToAdd, parsedRecords.Last()))
                    {
                        SetGroupIdForRows(recordToAdd, parsedRecords.Last());
                    }

                    parsedRecords.Add(recordToAdd);
                }
            }
        }

        private void SetGroupIdForRows(ParsedRecord current, ParsedRecord previous)
        {
            if (previous.GroupId == null)
            {
                Guid id = Guid.NewGuid();
                previous.GroupId = id;
                current.GroupId = id;
            }
            else
            {
                current.GroupId = previous.GroupId;
            }
        }

        private IRow FindHeaderRow()
        {
            try
            {
                IEnumerable<IRow> rows = Enumerable.Range(location.MinRow, location.MaxRow - location.MinRow + 1)
                    .Select(rowIndex => sheet.GetRow(rowIndex));

                return rows.First(row => headerRowResolver.IsHeaderRow(row, definition));
            }
            catch (InvalidOperationException)
            {
                exceptionManager.Register(new TableParserException($"No table header found for anchor: {definition.Anchor} and reqCols: {definition.RequiredColumns}"),
                    new ExceptionManager.Location(sheet.SheetName, location.MinRow + 1));
                return null;
            }
        }

        private Dictionary<string, int> MapAllHeaders(IRow headerRow)
        {
            Dictionary<string, int> allHeadersMapping = new Dictionary<string, int>();

            foreach (ICell cell in headerRow.Cells)
            {
                ICell mergedCell = mergedCellsResolver[cell.RowIndex, cell.ColumnIndex];

                if (mergedCell != null && ShouldNotBeIgnored(mergedCell))
                {
                    allHeadersMapping[$"{mergedCell}_{cell.ColumnIndex + 1}"] = cell.ColumnIndex;
                }
                else if (cell.CellType == CellType.String && cell.StringCellValue.Trim() != "" && ShouldNotBeIgnored(cell))
                {
                    allHeadersMapping[cell.StringCellValue] = cell.ColumnIndex;
                }
            }

            return allHeadersMapping;
        }

        private bool ShouldNotBeIgnored(ICell cell)
        {
            return !definition.IgnoredColumns.Any(column => column.Matches(cell));
        }

        private Metadata EnrichMetadata()
        {
            if (definition.Anchor != null)
            {
                return metadata.Plus(new KeyValuePair<string, object>(Metadata.ANCHOR, definition.Anchor));
            }
            else
            {
                return metadata;
            }
        }

        private bool IsExtractableRow(IRow row)
        {
            if (row == null)
                return false;

            return PrefilterCells(row).Any() && !headerRowResolver.IsHeaderRow(row, definition);
        }

        private bool IsDivider(IRow row, TableLocationWithHeader tableLocationWithHeader)
        {
            return FilteredCells(row, tableLocationWithHeader).Count() == 1;
        }

        private bool IsAllowedDivider(IRow row, TableLocationWithHeader tableLocationWithHeader, ISet<AbstractHeaderCell> allowedDividers)
        {
            return IsDivider(row, tableLocationWithHeader) && (allowedDividers.Count == 0 || allowedDividers.Any(allowedDivider => allowedDivider.Matches(FirstFilteredCell(row, tableLocationWithHeader))));
        }

        private ICell FirstFilteredCell(IRow row, TableLocationWithHeader tableLocationWithHeader)
        {
            return FilteredCells(row, tableLocationWithHeader).Single();
        }

        private IEnumerable<ICell> FilteredCells(IRow row, TableLocationWithHeader tableLocationWithHeader)
        {
            return PrefilterCells(row).Where(cell => tableLocationWithHeader.ColIndices.Contains(cell.ColumnIndex));
        }

        private IEnumerable<ICell> PrefilterCells(IRow row)
        {
            return row.Cells.Where(cell => cell.CellType != CellType.Blank && cell.ToString().Trim() != "");
        }
    }

    internal class TableLocation
    {
        public int MinRow { get; }
        public int MaxRow { get; }

        public TableLocation(int minRow, int maxRow)
        {
            MinRow = minRow;
            MaxRow = maxRow;
        }
    }

    internal class TableLocationWithHeader
    {
        public int MinRow { get; }
        public int HeaderRow { get; }
        public int MaxRow { get; }
        public ICollection<int> ColIndices { get; }

        public TableLocationWithHeader(int minRow, int headerRow, int maxRow, ICollection<int> colIndices)
        {
            if (minRow > maxRow)
                throw new ArgumentException("minRow should be less than or equal to maxRow");
            if (headerRow < minRow || headerRow > maxRow)
                throw new ArgumentException("headerRow should be between minRow and maxRow");

            MinRow = minRow;
            HeaderRow = headerRow;
            MaxRow = maxRow;
            ColIndices = colIndices;
        }

        public IEnumerable<int> Range()
        {
            return Enumerable.Range(HeaderRow, MaxRow - HeaderRow + 1);
        }
    }
}
