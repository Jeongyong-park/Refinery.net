using System;
using System.Collections.Generic;
using System.Linq;
using Refinery.Cell;
using Refinery.Exceptions;
using Refinery.Result;

namespace Refinery.Domain
{
    public class TableParserDefinition
    {
        public HashSet<AbstractHeaderCell> RequiredColumns { get; }
        public HashSet<AbstractHeaderCell> OptionalColumns { get; }
        public Func<RowParserData, ExceptionManager, RowParser> RowParserFactory { get; }
        public string Anchor { get; }
        public bool HasDivider { get; }
        public HashSet<AbstractHeaderCell> IgnoredColumns { get; }
        public HashSet<AbstractHeaderCell> AllowedDividers { get; }

        public TableParserDefinition(
            HashSet<AbstractHeaderCell> requiredColumns,
            HashSet<AbstractHeaderCell> optionalColumns ,
            Func<RowParserData, ExceptionManager, RowParser> rowParserFactory  =null,
            string anchor = null,
            bool hasDivider = false,
            HashSet<AbstractHeaderCell> ignoredColumns = null,
            HashSet<AbstractHeaderCell> allowedDividers = null
        )
        {
            RequiredColumns = requiredColumns ?? new HashSet<AbstractHeaderCell>();
            OptionalColumns = optionalColumns ?? new HashSet<AbstractHeaderCell>();
            RowParserFactory = rowParserFactory ?? ((rowParserData, exceptionManager) => new GenericRowParser(rowParserData, exceptionManager));
            Anchor = anchor;
            HasDivider = hasDivider;
            IgnoredColumns = ignoredColumns ?? new HashSet<AbstractHeaderCell>();
            AllowedDividers = allowedDividers ?? new HashSet<AbstractHeaderCell>();
        }

        public List<AbstractHeaderCell> AllColumns()
        {
            var allColumns = new HashSet<AbstractHeaderCell>();
            allColumns.UnionWith(RequiredColumns);
            allColumns.UnionWith(OptionalColumns);
            return allColumns.ToList();
        }
    }
}