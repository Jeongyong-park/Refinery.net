using NPOI.SS.UserModel;
using Refinery.Exceptions;
using Refinery.Result;
using Refinery.Tests.Configuration;

namespace Refinery.Tests;

public class TestGenericRowParser
{
    [Fact]
    public void TestGenericRowParserParsesTableAndProvidesCorrectRawParsedOutputIncludingUncapturedHeaders()
    {
        // given
        var file = new FileInfo("spreadsheet_examples/test_spreadsheet_uncaptured.xlsx");
        var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
        var workbook = WorkbookFactory.Create(stream);
        var fileName = file.Name;
        var exceptionManager = new ExceptionManager();

        // when

        var parsedRecords = new WorkbookParser(
            TestConfiguration.TestDefinition,
            workbook,
            exceptionManager,
            fileName
        ).Parse();

        // then
        Assert.Equal(1, exceptionManager.Exceptions.Count);


        Assert.Equal(new GenericParsedRecord(
                new Dictionary<string, object>
                {
                    ["workbook_name"] = fileName,
                    ["spreadsheet_name"] = "Sheet1",
                    ["divider"] = "",
                    ["number"] = 1d,
                    ["string"] = "one",
                    ["date"] = new DateTime(2021, 1, 1, 0, 0, 0),
                    ["optional_str"] = "exist",
                    ["row_number"] = 2f,
                    ["uncaptured"] = "Bob",
                    ["double"] = 1.05
                }),
                parsedRecords[0]);

    
        //Assert.Collection(parsedRecords,
        //    record1 => Assert.Equal(new GenericParsedRecord(
        //        new Dictionary<string, object?>
        //        {
        //            ["workbook_name"] = fileName,
        //            ["spreadsheet_name"] = "Sheet1",
        //            ["divider"] = "",
        //            ["number"] = 1d,
        //            ["string"] = "one",
        //            ["date"] = new DateTime(2021, 1, 1, 0, 0, 0),
        //            ["optional_str"] = "exist",
        //            ["row_number"] = 2f,
        //            ["uncaptured"] = "Bob",
        //            ["double"] = 1.05
        //        }), record1),
        //    record2 => Assert.Equal(new GenericParsedRecord(
        //        new Dictionary<string, object?>
        //        {
        //            ["workbook_name"] = fileName,
        //            ["spreadsheet_name"] = "Sheet1",
        //            ["string"] = "two",
        //            ["divider"] = "",
        //            ["number"] = 2d,
        //            ["date"] = new DateTime(2021, 1, 2, 0, 0, 0),
        //            ["row_number"] = 3f,
        //            ["uncaptured"] = "Alice",
        //            ["double"] = 2.5
        //        }), record2),
        //    record3 => Assert.Equal(new GenericParsedRecord(
        //        new Dictionary<string, object?>
        //        {
        //            ["workbook_name"] = fileName,
        //            ["spreadsheet_name"] = "Sheet1",
        //            ["divider"] = "",
        //            ["string"] = "three",
        //            ["number"] = 3d,
        //            ["date"] = new DateTime(2021, 1, 3, 0, 0, 0),
        //            ["row_number"] = 4f,
        //            ["double"] = 3.14
        //        }), record3)
        //);
    }
}