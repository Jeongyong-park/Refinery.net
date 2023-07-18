using NPOI.SS.UserModel;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Text;

namespace Refinery.Cell
{
    public class MergedCellsResolver
    {
        private readonly Dictionary<CellLocation, ICell> mergedCellsMap;

        public MergedCellsResolver(ISheet sheet)
        {
            var tempMap = new Dictionary<CellLocation, ICell>();

            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                var cellRangeAddress = sheet.GetMergedRegion(i);
                var firstRow = cellRangeAddress.FirstRow;
                var firstColumn = cellRangeAddress.FirstColumn;

                var cell = sheet.GetRow(firstRow)?.GetCell(firstColumn, MissingCellPolicy.RETURN_BLANK_AS_NULL);

                if (cell != null)
                {
                    for (int row = cellRangeAddress.FirstRow; row <= cellRangeAddress.LastRow; row++)
                    {
                        for (int col = cellRangeAddress.FirstColumn; col <= cellRangeAddress.LastColumn; col++)
                        {
                            tempMap[new CellLocation(row, col)] = cell;
                        }
                    }
                }
            }

            mergedCellsMap = tempMap;
        }

        public ICell this[int rowIndex, int columnIndex]
        {
            get { return mergedCellsMap.TryGetValue(new CellLocation(rowIndex, columnIndex), out var cell) ? cell : null; }
        }

        public class CellLocation
        {
            public int RowIndex { get; }
            public int ColumnIndex { get; }

            public CellLocation(int rowIndex, int columnIndex)
            {
                if (rowIndex < 0)
                    throw new ArgumentException("Row index should be non-negative.", nameof(rowIndex));

                if (columnIndex < 0)
                    throw new ArgumentException("Column index should be non-negative.", nameof(columnIndex));

                RowIndex = rowIndex;
                ColumnIndex = columnIndex;
            }

            public override int GetHashCode()
            {
                return RowIndex.GetHashCode() + ColumnIndex.GetHashCode();
            }



            public override bool Equals(object obj)
            {
                CellLocation o = obj as CellLocation;

                return o != null && (o.RowIndex == this.RowIndex || o.ColumnIndex == this.ColumnIndex);

            }


            public override String ToString()
            {
                return $"Row: {RowIndex}, Col:{ColumnIndex}";
            }
        }
    }
}
