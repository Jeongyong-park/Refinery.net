using System;
using System.Collections.Generic;

namespace Refinery.Exceptions
{
    public enum Level
    {
        CRITICAL,
        WARNING
    }

    public abstract class ManagedException : Exception
    {
        public Level Level { get; }

        public ManagedException(string message, Level level) : base(message)
        {
            Level = level;
        }

        public Dictionary<string, object> ExtractData()
        {
            return new Dictionary<string, object>
        {
            { "class", GetType().Name },
            { "level", Level },
            { "message", Message }
        };
        }

        public override bool Equals(object obj)
        {
            if (this == obj) return true;
            if (GetType() != obj?.GetType()) return false;

            var other = (ManagedException)obj;

            if (Message != other.Message) return false;
            if (Level != other.Level) return false;

            return true;
        }

        public override int GetHashCode()
        {
            int result = Message.GetHashCode();
            result = 31 * result + Level.GetHashCode();
            return result;
        }
    }

    public class CellParserException : ManagedException
    {
        public CellParserException(string message) : base(message, Level.WARNING)
        {
        }
    }

    public class TableParserException : ManagedException
    {
        public TableParserException(string message) : base(message, Level.WARNING)
        {
        }
    }

    public class SheetParserException : ManagedException
    {
        public SheetParserException(string message) : base(message, Level.WARNING)
        {
        }
    }

    public class WorkbookParserException : ManagedException
    {
        public WorkbookParserException(string message) : base(message, Level.CRITICAL)
        {
        }
    }

    public class UncategorizedException : ManagedException
    {
        public UncategorizedException(string message) : base(message, Level.CRITICAL)
        {
        }
    }
}