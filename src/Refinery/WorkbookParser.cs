using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;
using Refinery.Domain;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery
{
    public class WorkbookParser
    {
        private readonly WorkbookParserDefinition definition;
        private readonly IWorkbook workbook;
        private readonly ExceptionManager exceptionManager;
        private readonly string workbookName;

        public WorkbookParser(
            WorkbookParserDefinition definition,
            IWorkbook workbook,
            ExceptionManager exceptionManager = null,
            string workbookName = null)
        {
            this.definition = definition;
            this.workbook = workbook;
            this.exceptionManager = exceptionManager;
            this.workbookName = workbookName;
        }

        public List<ParsedRecord> Parse()
        {
            try
            {
                var sheetCount = workbook.NumberOfSheets;

                var parsedRecords = Enumerable.Range(0, sheetCount)
                    .Select(index => workbook.GetSheetAt(index))
                    .Where(sheet =>
                        definition.IncludeHidden || !workbook.IsSheetHidden(workbook.GetSheetIndex(sheet.SheetName)))
                    .Select(ResolveSheetParser)
                    .Where(parser => parser != null)
                    .SelectMany(parser => parser.Parse()).ToList();

                return parsedRecords;

            }
            catch (Exception e)
            {
                exceptionManager.Register(e);
                return new List<ParsedRecord>();
            }
        }

        private SheetParser ResolveSheetParser(ISheet sheet)
        {
            return definition.SpreadsheetParserDefinitions
                .Where(def => def.SheetNameFilter(sheet.SheetName))
                .Select(def => new SheetParser(def, sheet, exceptionManager, workbookName))
                .FirstOrDefault() ?? null;
        }
    }
}