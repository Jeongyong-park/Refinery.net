using System;
using System.Collections.Generic;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using Refinery.Domain;

namespace Refinery.Cell
{
    public class HeaderRowResolver
    {
        private MergedCellsResolver mergedCellsResolver;

        public HeaderRowResolver(MergedCellsResolver mergedCellsResolver)
        {
            this.mergedCellsResolver = mergedCellsResolver;
        }

        public Dictionary<AbstractHeaderCell, int> ResolveHeaderCellIndex(IRow headerRow, List<AbstractHeaderCell> headerCells)
        {
            var partitions = PartitionByType<OrderedHeaderCell, AbstractHeaderCell>(headerCells);
            var orderedCells = partitions.Item1;
            var unorderedCells = partitions.Item2;


            var result = ResolveOrderedHeaders(headerRow, orderedCells).Concat(ResolveUnorderedHeaders(headerRow, unorderedCells))
                .ToDictionary(entry => entry.Key, entry => entry.Value);

            return result.SelectMany(item =>
            {
                if (item.Key is MergedHeaderCell mergedCell)
                {
                    return mergedCell.HeaderCells.Select((hc, i) => new KeyValuePair<AbstractHeaderCell, int>(hc, item.Value + i));
                }
                else
                {
                    return new[] { item };
                }
            }).ToDictionary(x => x.Key, x => x.Value);

        }

        public bool IsHeaderRow(IRow row, TableParserDefinition definition)
        {
            var cellValues = AsMappedSet(row);
            var missingColumns = definition.RequiredColumns.Where(column => !column.Inside(cellValues));
            return !missingColumns.Any();
        }

        private IEnumerable<ICell> AsMappedSet(IRow row)
        {
            var result = row.Cells.Select(cell =>
            {
                ICell mergedCell = mergedCellsResolver[cell.RowIndex, cell.ColumnIndex];
                return (mergedCell != null && mergedCell.ColumnIndex == cell.ColumnIndex) ? mergedCell : cell;
            });

            return result;
        }

        private Dictionary<AbstractHeaderCell, int> ResolveOrderedHeaders(
            IRow row,
            List<OrderedHeaderCell> orderedCells
        )
        {
            //val matches = mutableMapOf<AbstractHeaderCell, Int>()
            //val sortedCells = orderedCells.sortedBy { it.priority }
            //row.asMappedSequence().forEach {
            //    cell->
            //        val filtered = sortedCells.filterNot { matches.contains(it.headerCell) }
            //    val headerCellOrNull = filtered.firstOrNull { oc->oc.matches(cell) }
            //    if (headerCellOrNull != null) matches[headerCellOrNull.headerCell] = cell.columnIndex
            //}
            //return matches
            var matches = new Dictionary<AbstractHeaderCell, int>();
            var sortedCells = orderedCells.OrderBy(c => c.Priority);

            AsMappedSet(row).ToList().ForEach(cell =>
            {
                var filtered = sortedCells.Where(it => !matches.ContainsKey(it.HeaderCell));
                var headerCellOrNull = filtered.FirstOrDefault(oc => oc.Matches((cell)));
                if (headerCellOrNull != null) matches[headerCellOrNull.HeaderCell] = cell.ColumnIndex;

            });

            return matches;
        }

        private Dictionary<AbstractHeaderCell, int> ResolveUnorderedHeaders(
            IRow row,
            List<AbstractHeaderCell> unorderedCells
        )
        {

            //private fun resolveUnorderedHeaders(
            //    row: Row,
            //    unorderedCells: List<AbstractHeaderCell>
            //  ): Map<AbstractHeaderCell, Int> {
            //    return row.asMappedSequence().mapNotNull {
            //        cell->
            //            val headerCellOrNull = unorderedCells.firstOrNull { hc->hc.matches(cell) }
            //        if (headerCellOrNull != null) Pair(headerCellOrNull, cell.columnIndex) else null
            //    }.toMap()
            //}
            var matches = new Dictionary<AbstractHeaderCell, int>();

            foreach (var cell in AsMappedSet(row))
            {
                var headerCellOrNull = unorderedCells.FirstOrDefault(hc => hc.Matches(cell));
                if (headerCellOrNull != null)
                {
                    matches.Add(headerCellOrNull, cell.ColumnIndex);
                }
            }

            return matches;
        }

        private Tuple<List<U>, List<T>> PartitionByType<U, T>(List<T> source)
        {
            var orderedCells = new List<U>();
            var unorderedCells = new List<T>();

            foreach (var element in source)
            {
                if (element is U u) orderedCells.Add(u);
                else unorderedCells.Add(element);
            }

            return Tuple.Create(orderedCells, unorderedCells);
        }
    }
}
