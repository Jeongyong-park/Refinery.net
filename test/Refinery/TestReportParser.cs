﻿using NPOI.SS.UserModel;
using Refinery.Cell;
using Refinery.Domain;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery.Tests
{
    public class TestReportsParser
    {
        [Fact]
        public void TestParsingMultipleTablesWithAnchors()
        {
            // given
            var definition = new WorkbookParserDefinition(
                new List<SheetParserDefinition>
                {
                    new SheetParserDefinition(
                        sheetNameFilter: (name) => true,
                        tableDefinitions: new List<TableParserDefinition>
                        {
                            new TableParserDefinition(
                                requiredColumns: new HashSet<AbstractHeaderCell> { new StringHeaderCell("string"), new StringHeaderCell("number"), new StringHeaderCell("date") },
                                optionalColumns: new HashSet<AbstractHeaderCell> { new StringHeaderCell("optional_str") },
                                anchor: "table {0}"
                            )

                        }
                    )
                }
            );

            var fileName = "Resources/spreadsheet_examples/test_spreadsheet_multitable_anchors.xlsx";
            var file = new FileInfo(fileName);
            var exceptionManager = new ExceptionManager();
            var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);

            // when'
            List<ParsedRecord> records;
            using (var workbook = WorkbookFactory.Create(stream))
            {
                var parser = new WorkbookParser(definition, workbook, exceptionManager, fileName);
                records = parser.Parse();
            }

            // then
            Assert.Empty(exceptionManager.Exceptions);
            Assert.Equal(9, records.Count);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "divider", string.Empty },
                    { "string", "one" },
                    { "number", 1D },
                    { "date", new DateTime(2021, 1, 1, 0, 0, 0) },
                    { "optional_str", "exist" },
                    { "anchor", "table 1" },
                    { "row_number", 3 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "divider", string.Empty },
                    { "string", "four" },
                    { "number", 4D },
                    { "date", new DateTime(2021, 1, 4, 0, 0, 0) },
                    { "anchor", "table 2" },
                    { "row_number", 10 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "divider", string.Empty },
                    { "string", "three" },
                    { "number", 3D },
                    { "date", new DateTime(2021, 1, 3, 0, 0, 0) },
                    { "anchor", "table 3" },
                    { "row_number", 20 }
                }
            ), records);
        }

        [Fact]
        public void TestParsingMultipleTablesWithoutAnchorsAndWithBadFormatting()
        {
            // given
            var stringHeaderCell = new StringHeaderCell("string");
            var numberHeaderCell = new StringHeaderCell("number");
            var dateHeaderCell = new StringHeaderCell("date");
            var optionalStringHeaderCell = new StringHeaderCell("optional_str");

            var definition = new WorkbookParserDefinition(
                new List<SheetParserDefinition>
                {
                    new SheetParserDefinition(
                        sheetNameFilter: (name) => true,
                        tableDefinitions: new List<TableParserDefinition>
                        {
                            new TableParserDefinition(
                                new HashSet<AbstractHeaderCell> { stringHeaderCell, numberHeaderCell, dateHeaderCell },
                                new HashSet<AbstractHeaderCell> { optionalStringHeaderCell },
                                rowParserFactory:(a,b)=>new GenericRowParser(a,b)
                            )
                        }
                    )
                }
            );

            var fileName = "Resources/spreadsheet_examples/test_spreadsheet_multitable_unaligned.xlsx";
            var file = new FileInfo(fileName);
            var exceptionManager = new ExceptionManager();
            List<ParsedRecord> records;
            var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);

            // when
            using (var workbook = WorkbookFactory.Create(stream))
            {
                var parser = new WorkbookParser(definition, workbook, exceptionManager, fileName);
                records = parser.Parse();
            };

            // then
            Assert.Empty(exceptionManager.Exceptions);
            Assert.Equal(9, records.Count);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "one" },
                    { "number", 1 },
                    { "date", new DateTime(2021, 1, 1, 0, 0, 0) },
                    { "optional_str", "exist" },
                    { "row_number", 3 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "four" },
                    { "number", 4 },
                    { "date", new DateTime(2021, 1, 4, 0, 0, 0) },
                    { "row_number", 10 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "three" },
                    { "number", 3 },
                    { "date", new DateTime(2021, 1, 3, 0, 0, 0) },
                    { "row_number", 20 }
                }
            ), records);
        }

        [Fact]
        public void TestParsingMultipleSheetsWithSheetFilter()
        {
            // given
            var stringHeaderCell = new StringHeaderCell("string");
            var numberHeaderCell = new StringHeaderCell("number");
            var dateHeaderCell = new StringHeaderCell("date");
            var optionalStringHeaderCell = new StringHeaderCell("optional_str");

            var definition = new WorkbookParserDefinition(
                new List<SheetParserDefinition>
                {
                    new SheetParserDefinition(
                        sheetNameFilter: (name) => name != "Sheet3",
                        tableDefinitions: new List<TableParserDefinition>
                        {
                            new TableParserDefinition(
                                new HashSet<AbstractHeaderCell> { stringHeaderCell, numberHeaderCell, dateHeaderCell },
                                new HashSet<AbstractHeaderCell> { optionalStringHeaderCell },
                                rowParserFactory: (rpd,em)=>new GenericRowParser(rpd,em)
                            )
                        }
                    )
                }
            );

            var fileName = "Resources/spreadsheet_examples/test_spreadsheet_multisheet.xlsx";
            var file = new FileInfo(fileName);
            var exceptionManager = new ExceptionManager();
            var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
            // when
            List<ParsedRecord> records;
            using (var workbook = WorkbookFactory.Create(stream))
            {
                var parser = new WorkbookParser(definition, workbook, exceptionManager, fileName);
                records = parser.Parse();
            }

            // then
            Assert.Empty(exceptionManager.Exceptions);
            Assert.Equal(6, records.Count);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "one" },
                    { "number", 1 },
                    { "date", new DateTime(2021, 1, 1, 0, 0, 0) },
                    { "optional_str", "exist" },
                    { "row_number", 2 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet2" },
                    { "string", "four" },
                    { "number", 4 },
                    { "date", new DateTime(2021, 1, 4, 0, 0, 0) },
                    { "optional_str", "exist" },
                    { "row_number", 2 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet2" },
                    { "string", "six" },
                    { "number", 6 },
                    { "date", new DateTime(2021, 1, 6, 0, 0, 0) },
                    { "row_number", 4 }
                }
            ), records);
        }

        [Fact]
        public void TestParsingMergedRowsAndColumns()
        {
            // given
            var stringHeaderCell = new StringHeaderCell("string");
            var numberHeaderCell = new StringHeaderCell("number");
            var dateHeaderCell = new StringHeaderCell("date");
            var optionalStringHeaderCell = new StringHeaderCell("optional_str");
            var optionalStringHeaderCell2 = new StringHeaderCell("optional_str2");
            var mergedCols = new MergedHeaderCell(
                new StringHeaderCell("optional_str"),
                new HashSet<AbstractHeaderCell> { optionalStringHeaderCell, optionalStringHeaderCell2 }
            );

            var definition = new WorkbookParserDefinition(
                new List<SheetParserDefinition>
                {
                    new SheetParserDefinition(
                        sheetNameFilter: (name) => name == "Sheet1",
                        tableDefinitions: new List<TableParserDefinition>
                        {
                            new TableParserDefinition(
                                new HashSet<AbstractHeaderCell> { stringHeaderCell, numberHeaderCell, dateHeaderCell },
                                new HashSet<AbstractHeaderCell> { mergedCols },
                                rowParserFactory: (rpd,em)=>new GenericRowParser(rpd,em)
                            )
                        }
                    )
                }
            );

            var fileName = "Resources/spreadsheet_examples/test_spreadsheet_merged_cells.xlsx";
            var file = new FileInfo(fileName);
            var exceptionManager = new ExceptionManager();
            var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
            List<ParsedRecord> records;
            using (var workbook = WorkbookFactory.Create(stream))
            {
                var parser = new WorkbookParser(definition, workbook, exceptionManager, fileName);
                records = parser.Parse();
            }
            // then
            Assert.Empty(exceptionManager.Exceptions);
            Assert.Equal(5, records.Count);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "one" },
                    { "number", 1 },
                    { "date", new DateTime(2021, 1, 1, 0, 0, 0) },
                    { "optional_str_4", "exist" },
                    { "optional_str_5", "exist2" },
                    { "row_number", 2 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "two" },
                    { "number", 2 },
                    { "date", new DateTime(2021, 1, 2, 0, 0, 0) },
                    { "optional_str_4", "exist" },
                    { "optional_str_5", "exist2" },
                    { "row_number", 3 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "three" },
                    { "number", 3 },
                    { "date", new DateTime(2021, 1, 3, 0, 0, 0) },
                    { "row_number", 4 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "four and five" },
                    { "number", 4 },
                    { "date", new DateTime(2021, 1, 4, 0, 0, 0) },
                    { "optional_str_4", "same" },
                    { "optional_str_5", "same" },
                    { "row_number", 5 }
                }
            ), records);
            Assert.Contains(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    { "workbook_name", fileName },
                    { "spreadsheet_name", "Sheet1" },
                    { "string", "four and five" },
                    { "number", 5 },
                    { "date", new DateTime(2021, 1, 5, 0, 0, 0) },
                    { "row_number", 6 }
                }
            ), records);
        }
    }
}
