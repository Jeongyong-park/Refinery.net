﻿using NPOI.HSSF.Record;
using NPOI.SS.UserModel;
using Refinery.Cell;
using Refinery.Domain;
using Refinery.Exceptions;
using Refinery.Result;
using Refinery.Tests.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Refinery.Tests
{
    public class TestMetadataParser
    {
        [Fact]
        public void Test()
        {
            // given
            var file = new FileInfo("Resources/spreadsheet_examples/test_spreadsheet_metadata.xlsx");
            var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
            var workbook = WorkbookFactory.Create(stream);
            var fileName = file.Name;
            var exceptionManager = new ExceptionManager();

            // when

            var parsedRecords = new WorkbookParser(
                TestWithMetadataDefinition,
                workbook,
                exceptionManager,
                fileName
            ).Parse();


            // then
            Assert.Contains(
            new GenericParsedRecord(
             new Dictionary<string, object>
             {
                 ["workbook_name"] = fileName,
                 ["spreadsheet_name"] = "Sheet1",
                 ["key1"] = "value1",
                 ["key2"] = "value2",
                 ["key3"] = "value3",
                 ["key4"] = "value4",
                 ["key5"] = "value5",
                 ["key6"] = "value6",

                 ["row_number"] = 4,

                 ["string"] = "one",
                 ["number"] = 1,
                 ["date"] = new DateTime(2021, 1, 1, 0, 0, 0),
                 ["optional_str"] = "exist",


             }),
              parsedRecords);

            Assert.Contains(
            new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    ["workbook_name"] = fileName,
                    ["spreadsheet_name"] = "Sheet1",
                    ["key1"] = "value1",
                    ["key2"] = "value2",
                    ["key3"] = "value3",
                    ["key4"] = "value4",
                    ["key5"] = "value5",
                    ["key6"] = "value6",

                    ["row_number"] = 5,

                    ["string"] = "two",
                    ["number"] = 2,
                    ["date"] = new DateTime(2021, 1, 2, 0, 0, 0),

                }),
            parsedRecords);

            
            Assert.Contains(
            new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    ["workbook_name"] = fileName,
                    ["spreadsheet_name"] = "Sheet1",
                    ["key1"] = "value1",
                    ["key2"] = "value2",
                    ["key3"] = "value3",
                    ["key4"] = "value4",
                    ["key5"] = "value5",
                    ["key6"] = "value6",

                    ["row_number"] = 6,

                    ["string"] = "three",
                    ["number"] = 3,
                    ["date"] = new DateTime(2021, 1, 3, 0, 0, 0),

                }),
            parsedRecords);

            var targetRecord = parsedRecords.Where(item => item.ExtractedRawData["number"].Equals(222d)).FirstOrDefault();

            Assert.NotNull(targetRecord);
            Assert.Equal("value1-2", targetRecord.ExtractedRawData["key1"].ToString());
            Assert.Equal("value6-2", targetRecord.ExtractedRawData["key6"].ToString());

        }


        public static WorkbookParserDefinition TestWithMetadataDefinition = new WorkbookParserDefinition
        {
            SpreadsheetParserDefinitions = new List<SheetParserDefinition>
            {
                new SheetParserDefinition(
                    name=>true,
                    tableDefinitions: new List<TableParserDefinition>{
                        new TableParserDefinition(
                            new HashSet<AbstractHeaderCell>{
                                new SimpleHeaderCell("string"),
                                new SimpleHeaderCell("number"),
                                new SimpleHeaderCell("date"),
                            },
                            new HashSet<AbstractHeaderCell>{
                                new SimpleHeaderCell("optional_str"),
                            }
                        )
                    },
                    metadataParserDefinition: new List<MetadataEntryDefinition>{
                        new MetadataEntryDefinition("key1", "key1", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),
                        new MetadataEntryDefinition("key2", "key2", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),
                        new MetadataEntryDefinition("key3", "key3", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),

                        new MetadataEntryDefinition("key4", "key4", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),
                        new MetadataEntryDefinition("key5", "key5", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),
                        new MetadataEntryDefinition("key6", "key6", MetadataValueLocation.NEXT_CELL_VALUE, (cell)=>cell.ToString()),
                        
                    }
                )
            }
        };
    }
}
