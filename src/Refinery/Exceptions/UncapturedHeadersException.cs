using System;
using System.Collections.Generic;
using System.Linq;

namespace Refinery.Exceptions
{
    public class UncapturedHeadersException : ManagedException
    {
        public List<UncapturedHeaderCell> UncapturedHeaders { get; }

        public UncapturedHeadersException(List<UncapturedHeaderCell> uncapturedHeaders)
            : base(GenerateMessage(uncapturedHeaders), Level.WARNING)
        {
            if (uncapturedHeaders.Count == 0)
                throw new ArgumentException("Should be at least 1 uncaptured header");

            UncapturedHeaders = uncapturedHeaders;
        }

        private static string GenerateMessage(List<UncapturedHeaderCell> uncapturedHeaders)
        {
            return string.Join(" ", uncapturedHeaders.Select(cell => $"{cell.Name} @ {cell.Index + 1}"));
        }

        public class UncapturedHeaderCell
        {
            public string Name { get; }
            public int Index { get; }

            public UncapturedHeaderCell(string name, int index)
            {
                if (index < 0)
                    throw new ArgumentException("Index should be non-negative");

                Name = name;
                Index = index;
            }
        }
    }
}