﻿using System.Collections;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using Refinery.Domain;
using Refinery.Result;

namespace Refinery
{
    internal class MetadataParser
    {
        private readonly List<MetadataEntryDefinition> definitions;
        private readonly ISheet sheet;
        private readonly string workbookName;

        public MetadataParser(List<MetadataEntryDefinition> definitions, ISheet sheet, string workbookName)
        {
            this.definitions = definitions;
            this.sheet = sheet;
            this.workbookName = workbookName;
        }

        public Metadata ExtractMetadata()
        {
            var metadata = new Dictionary<string, object>
            {
                { Metadata.SPREADSHEET_NAME, sheet.SheetName }
            };
            if (workbookName != null)
                metadata[Metadata.WORKBOOK_NAME] = workbookName;

            foreach (var definition in definitions)
            {
                var kv = FindMetadata(definition);
                metadata[kv.Key] = kv.Value;
            }

            return new Metadata(metadata);
        }

        private KeyValuePair<string, object> FindMetadata(MetadataEntryDefinition definition)
        {

            ICell matchingCell = null;
            IEnumerator enumerator = sheet.GetRowEnumerator();

            while (enumerator.MoveNext())
            {
                IRow row = enumerator.Current as IRow;
                if (row != null)
                {
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
                }
            }
            if (matchingCell == null) return new KeyValuePair<string, object>();

            ICell cell;
            switch (definition.ValueLocation)
            {
                case MetadataValueLocation.SameCellValue:
                    cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex);
                    break;
                case MetadataValueLocation.NextCellValue:
                    cell = sheet.GetRow(matchingCell.RowIndex).GetCell(matchingCell.ColumnIndex + 1);
                    break;
                default:
                    cell = (ICell)null;
                    break;
            }

            return new KeyValuePair<string, object>(definition.MetadataName, definition.Extractor.Invoke(cell));
        }
    }
}