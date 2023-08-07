using System;
using System.Collections;
using System.Collections.Generic;
using NPOI.SS.Formula.UDF;
using NPOI.SS.UserModel;
using Refinery.Cell;
using Refinery.Domain;
using Refinery.Result;

namespace Refinery
{
    internal class MetadataParser
    {
        private readonly List<MetadataEntryDefinition> definitions;
        private readonly ISheet sheet;
        private readonly string workbookName;
        private readonly MergedCellsResolver mergedCellsResolver;

        public MetadataParser(List<MetadataEntryDefinition> definitions, ISheet sheet, string workbookName, MergedCellsResolver mergedCellsResolver)
        {
            this.definitions = definitions;
            this.sheet = sheet;
            this.workbookName = workbookName;
            this.mergedCellsResolver = mergedCellsResolver;
        }

        public (Metadata, MetadataLocation) ExtractMetadata(int startRowNum = 0)
        {
            int minRowNum = int.MaxValue, maxRowNum = int.MinValue;
            var metadata = new Dictionary<string, object>
            {
                { Metadata.SPREADSHEET_NAME, sheet.SheetName }
            };
            if (workbookName != null)
                metadata[Metadata.WORKBOOK_NAME] = workbookName;


            foreach (var definition in definitions)
            {
                var (kv, rowNum) = FindMetadata(definition, startRowNum);
                if (kv.Key != null)
                {
                    metadata[kv.Key] = kv.Value;
                }
                minRowNum = Math.Min(rowNum, minRowNum);
                maxRowNum = Math.Max(rowNum, maxRowNum);
            }

            return (new Metadata(metadata), new MetadataLocation(minRowNum, maxRowNum));
        }

        private (KeyValuePair<string, object>, int) FindMetadata(MetadataEntryDefinition definition, int startRowNum)
        {

            ICell matchingCell = null;
            IEnumerator enumerator = sheet.GetRowEnumerator();

            while (enumerator.MoveNext())
            {
                IRow row = enumerator.Current as IRow;
                if (row != null)
                {

                    if (row.RowNum < startRowNum)
                    {
                        continue;
                    }

                    IEnumerator<ICell> cellEnumerator = row.GetEnumerator();
                    while (cellEnumerator.MoveNext())
                    {
                        var currentCell = cellEnumerator.Current;
                        if (currentCell.ToString().Contains(definition.MatchingCellKey))
                        {
                            matchingCell = currentCell;
                            break;
                        }
                    }
                    if (matchingCell != null)
                    {
                        break;
                    }
                }
            }

            if (matchingCell == null) return (new KeyValuePair<string, object>(), -1);

            ICell cell;
            switch (definition.ValueLocation)
            {
                case MetadataValueLocation.PREVIOUS_ROW_VALUE:
                    cell = sheet.GetRow(matchingCell.RowIndex - 1).GetCell(matchingCell.ColumnIndex);
                    break;

                case MetadataValueLocation.NEXT_ROW_VALUE:
                    cell = sheet.GetRow(matchingCell.RowIndex + 1).GetCell(matchingCell.ColumnIndex);
                    break;

                case MetadataValueLocation.PREVIOUS_CELL_VALUE:
                    cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex - 1);
                    break;

                case MetadataValueLocation.SAME_CELL_VALUE:
                    cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex);
                    break;

                case MetadataValueLocation.NEXT_CELL_VALUE:
                    var mergedCell = mergedCellsResolver[matchingCell.RowIndex, matchingCell.ColumnIndex];
                    if (mergedCell?.IsMergedCell ?? false)
                    {
                        int idx = 1;
                        var keyCellValue = mergedCell.StringCellValue;

                        while (
                            (mergedCell.Row.LastCellNum >= (matchingCell.ColumnIndex + idx)) &&
                            (mergedCellsResolver[matchingCell.RowIndex, matchingCell.ColumnIndex + idx]?.ToString() == keyCellValue.ToString())
                        )
                        {
                            idx++;
                        }
                        cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex + idx);
                        break;
                    }
                    cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex + 1);
                    break;
                default:
                    cell = (ICell)null;
                    break;
            }

            return (new KeyValuePair<string, object>(definition.MetadataName, definition.Extractor.Invoke(cell)), cell?.RowIndex ?? -1);
        }
        public class MetadataLocation
        {
            public int MinRow { get; }
            public int MaxRow { get; }

            public MetadataLocation(int minRow, int maxRow)
            {
                MinRow = minRow;
                MaxRow = maxRow;
            }
        }
    }
}