using System.Globalization;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Refinery.Tests.Cell.Parser
{
    public class CellParserTest : IDisposable
    {
        private readonly IWorkbook workbook;
        private readonly ISheet sheet;

        protected CellParserTest()
        {
            NPOI.Util.LocaleUtil.SetUserLocale(CultureInfo.GetCultureInfo("en-US"));

            var file = new FileInfo("spreadsheet_examples/cell_parsers.xlsx");
            using (var stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                workbook = new XSSFWorkbook(stream);
                
                sheet = workbook.GetSheet("cell_parsers");
            }
        }

        public void Dispose() => workbook.Close();

        ICell GetCell(int i, int j) => sheet.GetRow(i).GetCell(j);


        //[Fact]
        public ICell StringCell()
        {
            var cell = GetCell(1, 0);
            //Assert.Equal(CellType.String, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell EmptyStringCell()
        {
            var cell = GetCell(1, 1);
            //Assert.Equal(CellType.String, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DoubleCell()
        {
            var cell = GetCell(1, 2);
            //Assert.Equal(CellType.Numeric, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell IntCell()
        {
            var cell = GetCell(1, 3);
            //Assert.Equal(CellType.Numeric, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateCell()
        {
            var cell = GetCell(1, 4);
            //Assert.Equal(CellType.Numeric, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell BoolCell()
        {
            var cell = GetCell(1, 5);
            //Assert.Equal(CellType.Boolean, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell NullCell()
        {
            var cell = GetCell(1, 6);
            //Assert.Null(cell);
            return cell;
        }

        //[Fact]
        public ICell DoubleAsStringCell()
        {
            var cell = GetCell(1, 7);
            //Assert.Equal(CellType.String, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DoubleInt()
        {
            var cell = GetCell(1, 8);
            //Assert.Equal(CellType.Numeric, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateTime()
        {
            var cell = GetCell(1, 9);
            //Assert.Equal(CellType.Numeric, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateStr()
        {
            var cell = GetCell(1, 10);
            //Assert.Equal(CellType.String, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateTimeStr()
        {
            var cell = GetCell(1, 11);
            //Assert.Equal(CellType.String, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell StringCellFromFormula()
        {
            var cell = GetCell(2, 0);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell EmptyStringCellFromFormula()
        {
            var cell = GetCell(2, 1);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DoubleCellFromFormula()
        {
            var cell = GetCell(2, 2);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell IntCellFromFormula()
        {
            var cell = GetCell(2, 3);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateCellFromFormula()
        {
            var cell = GetCell(2, 4);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell BoolCellFromFormula()
        {
            var cell = GetCell(2, 5);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DoubleAsStringCellFromFormula()
        {
            var cell = GetCell(2, 7);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DoubleIntFromFormula()
        {
            var cell = GetCell(2, 8);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateTimeFromFormula()
        {
            var cell = GetCell(2, 9);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateStrFromFormula()
        {
            var cell = GetCell(2, 10);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }

        //[Fact]
        public ICell DateTimeStrFromFormula()
        {
            var cell = GetCell(2, 11);
            //Assert.Equal(CellType.Formula, cell.CellType);
            return cell;
        }
    }
}
