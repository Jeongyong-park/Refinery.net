using NPOI.SS.UserModel;
using Refinery.Cell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Refinery.Tests.Cell
{
    public class TestHeaderRowResolver
    {
        private IWorkbook workbook;
        private ISheet sheet;

        private MergedCellsResolver mergedCellsResolver;
        private HeaderRowResolver headerRowResolver;

        public TestHeaderRowResolver()
        {
            var file = new FileInfo("Resources/spreadsheet_examples/header_cells.xlsx");

            using (var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                workbook = WorkbookFactory.Create(stream);
                sheet = workbook.GetSheet("header_cells");
                mergedCellsResolver = new MergedCellsResolver(sheet);
                headerRowResolver = new HeaderRowResolver(mergedCellsResolver);
            }
        }

        private IRow GetHeaderRow()
        {
            return sheet.GetRow(0);
        }

        private IRow GetComplexHeaderRow()
        {
            return sheet.GetRow(5);
        }

        [Fact]
        public void TestStringCellParserWithExactMatching()
        {
            // given
            IRow headerRow = GetHeaderRow();

            SimpleHeaderCell simple1 = new SimpleHeaderCell("simpleHeader1");
            SimpleHeaderCell simple2 = new SimpleHeaderCell("simpleHeader2");
            SimpleHeaderCell merged1 = new SimpleHeaderCell("mergedHeader1");
            SimpleHeaderCell merged2 = new SimpleHeaderCell("mergedHeader2");
            SimpleHeaderCell regex1 = new SimpleHeaderCell("regexHeader");
            SimpleHeaderCell regex2 = new SimpleHeaderCell("regexHeader2");

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(
                headerRow,
                new List<AbstractHeaderCell> { simple1, simple2, merged1, merged2, regex1, regex2 }
            );

            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { simple1, 0 },
                    { simple2, 1 },
                    { merged1, 2 },
                    { merged2, 4 },
                    { regex1, 6 },
                    { regex2, 7 }
                },
                result
            );
        }

        [Fact]
        public void TestRegexCellParser()
        {
            // given
            IRow headerRow = GetHeaderRow();

            RegexHeaderCell regex1 = new RegexHeaderCell("regexHeader$");
            RegexHeaderCell regex2 = new RegexHeaderCell("regexHeader2");

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(headerRow, new List<AbstractHeaderCell> { regex1, regex2 });
            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { regex1, 6 },
                    { regex2, 7 }
                },
                result
            );
        }

        [Fact]
        public void TestMergedCellParser()
        {
            // given
            IRow headerRow = GetHeaderRow();

            SimpleHeaderCell m1 = new SimpleHeaderCell("merged1");
            SimpleHeaderCell m2 = new SimpleHeaderCell("merged2");
            SimpleHeaderCell m3 = new SimpleHeaderCell("merged3");
            SimpleHeaderCell m4 = new SimpleHeaderCell("merged4");

            MergedHeaderCell merged1 = new MergedHeaderCell(new StringHeaderCell("mergedHeader1"), new HashSet<AbstractHeaderCell> { m1, m2 });
            MergedHeaderCell merged2 = new MergedHeaderCell(new StringHeaderCell("mergedHeader2"), new HashSet<AbstractHeaderCell> { m3, m4 });

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(headerRow, new List<AbstractHeaderCell> { merged1, merged2 });
            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { m1, 2 },
                    { m2, 3 },
                    { m3, 4 },
                    { m4, 5 }
                },
                result
            );
        }

        [Fact]
        public void TestOrderedCellParser()
        {
            // given
            IRow headerRow = GetHeaderRow();

            StringHeaderCell simple1 = new StringHeaderCell("header");
            StringHeaderCell simple2 = new StringHeaderCell("header");
            StringHeaderCell simple3 = new StringHeaderCell("header");
            StringHeaderCell simple4 = new StringHeaderCell("header");
            StringHeaderCell simple5 = new StringHeaderCell("header");
            StringHeaderCell simple6 = new StringHeaderCell("header");

            OrderedHeaderCell ordered1 = new OrderedHeaderCell(simple1, 6);
            OrderedHeaderCell ordered2 = new OrderedHeaderCell(simple2, 5);
            OrderedHeaderCell ordered3 = new OrderedHeaderCell(simple3, 4);
            OrderedHeaderCell ordered4 = new OrderedHeaderCell(simple4, 3);
            OrderedHeaderCell ordered5 = new OrderedHeaderCell(simple5, 2);
            OrderedHeaderCell ordered6 = new OrderedHeaderCell(simple6, 1);

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(
                headerRow,
                new List<AbstractHeaderCell> { ordered5, ordered3, ordered2, ordered4, ordered6, ordered1 }
            );
            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { simple6, 0 },
                    { simple5, 1 },
                    { simple4, 2 },
                    { simple3, 4 },
                    { simple2, 6 },
                    { simple1, 7 }
                },
                result
            );
        }

        [Fact]
        public void TestOrderedMergedCellParser()
        {
            // given
            IRow headerRow = GetHeaderRow();

            SimpleHeaderCell m1 = new SimpleHeaderCell("merged1");
            SimpleHeaderCell m2 = new SimpleHeaderCell("merged2");
            SimpleHeaderCell m3 = new SimpleHeaderCell("merged3");
            SimpleHeaderCell m4 = new SimpleHeaderCell("merged4");

            MergedHeaderCell merged1 = new MergedHeaderCell(new StringHeaderCell("merged"), new HashSet<AbstractHeaderCell> { m1, m2 });
            MergedHeaderCell merged2 = new MergedHeaderCell(new StringHeaderCell("merged"), new HashSet<AbstractHeaderCell> { m3, m4 });
            OrderedHeaderCell ordered1 = new OrderedHeaderCell(merged1, 1);
            OrderedHeaderCell ordered2 = new OrderedHeaderCell(merged2, 2);

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(headerRow, new List<AbstractHeaderCell> { ordered1, ordered2 });
            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { m1, 2 },
                    { m2, 3 },
                    { m3, 4 },
                    { m4, 5 }
                },
                result
            );
        }

        [Fact]
        public void TestComplexOrderedMergedCellParser()
        {
            // given
            IRow headerRow = GetComplexHeaderRow();

            SimpleHeaderCell s1 = new SimpleHeaderCell("simpleHeader1");
            SimpleHeaderCell m1 = new SimpleHeaderCell("merged1");
            SimpleHeaderCell m2 = new SimpleHeaderCell("merged2");
            SimpleHeaderCell m3 = new SimpleHeaderCell("merged3");
            SimpleHeaderCell m4 = new SimpleHeaderCell("merged4");
            SimpleHeaderCell r1 = new SimpleHeaderCell("regexHeader");
            SimpleHeaderCell r2 = new SimpleHeaderCell("regexHeader2");

            MergedHeaderCell merged1 = new MergedHeaderCell(new StringHeaderCell("merged"), new HashSet<AbstractHeaderCell> { m1, m2 });
            MergedHeaderCell merged2 = new MergedHeaderCell(new StringHeaderCell("merged"), new HashSet<AbstractHeaderCell> { m3, m4 });
            OrderedHeaderCell ordered1 = new OrderedHeaderCell(merged1, 1);
            OrderedHeaderCell ordered2 = new OrderedHeaderCell(merged2, 2);

            // when
            var result = headerRowResolver.ResolveHeaderCellIndex(
                headerRow,
                new List<AbstractHeaderCell> { s1, ordered1, ordered2, r1, r2 }
            );
            // then
            Assert.Equal(
                new Dictionary<AbstractHeaderCell, int>
                {
                    { s1, 0 },
                    { m1, 1 },
                    { m2, 2 },
                    { m3, 3 },
                    { m4, 4 },
                    { r1, 5 },
                    { r2, 6 }
                },
                result
            );
        }

        [Fact]
        public void TestMergedOrderedCellThrowsException()
        {
            Assert.Throws<ArgumentException>(() => new MergedHeaderCell(new OrderedHeaderCell(new SimpleHeaderCell(""), 1), new HashSet<AbstractHeaderCell>()));
        }
    }
}
