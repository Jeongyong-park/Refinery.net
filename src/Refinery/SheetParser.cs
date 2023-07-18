using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Refinery.Cell;
using Refinery.Domain;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery
{
    internal class SheetParser
    {
        private readonly SheetParserDefinition definition;
        private readonly ISheet sheet;
        private readonly ExceptionManager exceptionManager;
        private readonly string workbookName;
        private readonly MergedCellsResolver mergedCellsResolver;
        private readonly HeaderRowResolver headerRowResolver;

        public SheetParser(SheetParserDefinition definition, ISheet sheet, ExceptionManager exceptionManager, string workbookName)
        {
            this.definition = definition;
            this.sheet = sheet;
            this.exceptionManager = exceptionManager;
            this.workbookName = workbookName;
            this.mergedCellsResolver = new MergedCellsResolver(sheet);
            this.headerRowResolver = new HeaderRowResolver(mergedCellsResolver);
        }

        public List<ParsedRecord> Parse()
        {
            try
            {
                Metadata metadata = new MetadataParser(definition.MetadataParserDefinition, sheet, workbookName).ExtractMetadata();
                List<TableParser> tableParsers = ResolveTableParsers(metadata);
                List<ParsedRecord> parsedRecords = new List<ParsedRecord>();
                foreach (TableParser tableParser in tableParsers)
                {
                    parsedRecords.AddRange(tableParser.Parse());
                }
                return parsedRecords;
            }
            catch (ManagedException e)
            {
                exceptionManager.Register(e, new ExceptionManager.Location(sheet.SheetName));
                return new List<ParsedRecord>();
            }
        }

        private List<TableParser> ResolveTableParsers(Metadata metadata)
        {
            List<Tuple<TableParserDefinition, TableLocation>> tableLocations = ResolveTableLocations();
            List<TableParser> tableParsers = new List<TableParser>();
            foreach (Tuple<TableParserDefinition, TableLocation> tableLocation in tableLocations)
            {
                TableParser tableParser = new TableParser(sheet, tableLocation.Item1, metadata, tableLocation.Item2, mergedCellsResolver, exceptionManager, headerRowResolver);
                tableParsers.Add(tableParser);
            }
            return tableParsers;
        }

        private List<Tuple<TableParserDefinition, TableLocation>> ResolveTableLocations()
        {
            List<Tuple<TableParserDefinition, int>> definitionsBeginning = new List<Tuple<TableParserDefinition, int>>();
            IEnumerator<IRow> rowIterator = (IEnumerator<IRow>)sheet.GetRowEnumerator();
            while (rowIterator.MoveNext())
            {
                IRow row = rowIterator.Current;
                TableParserDefinition tableDefinition = TryMatchToTableParserDefinition(row);
                if (tableDefinition != null)
                {
                    definitionsBeginning.Add(Tuple.Create(tableDefinition, row.RowNum));
                }
            }
            definitionsBeginning.Sort((def1, def2) => def1.Item2.CompareTo(def2.Item2));
            if (definitionsBeginning.Count == 0)
            {
                throw new SheetParserException("Could not locate any tables");
            }

            List<Tuple<TableParserDefinition, TableLocation>> tableLocations = new List<Tuple<TableParserDefinition, TableLocation>>();
            for (int i = 0; i < definitionsBeginning.Count - 1; i++)
            {
                Tuple<TableParserDefinition, int> def1 = definitionsBeginning[i];
                Tuple<TableParserDefinition, int> def2 = definitionsBeginning[i + 1];
                TableLocation location = new TableLocation(def1.Item2, def2.Item2 - 1);
                tableLocations.Add(Tuple.Create(def1.Item1, location));
            }
            Tuple<TableParserDefinition, int> lastDefinition = definitionsBeginning[definitionsBeginning.Count - 1];
            TableLocation lastLocation = new TableLocation(lastDefinition.Item2, sheet.LastRowNum);
            tableLocations.Add(Tuple.Create(lastDefinition.Item1, lastLocation));

            return tableLocations;
        }

        private TableParserDefinition TryMatchToTableParserDefinition(IRow row)
        {
            foreach (TableParserDefinition sd in definition.TableDefinitions)
            {
                if (sd.Anchor != null)
                {
                    List<string> cellValues = new List<string>();
                    IEnumerator<ICell> cellIterator = row.GetEnumerator();
                    while (cellIterator.MoveNext())
                    {
                        ICell cell = cellIterator.Current;
                        if (cell.CellType == CellType.String)
                        {
                            cellValues.Add(cell.StringCellValue.Trim().ToLower());
                        }
                    }

                    var regexExpr = $"{sd.Anchor.ToLower().Replace("{0}", @"\d+")}";
                    if (cellValues.Exists(val => Regex.IsMatch(val, regexExpr)) && !headerRowResolver.IsHeaderRow(row, sd))
                    {
                        return sd;
                    }
                }
                else if (headerRowResolver.IsHeaderRow(row, sd))
                {
                    return sd;
                }
            }
            return null;
        }
    }
}