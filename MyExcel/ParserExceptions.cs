using System;

namespace MyExcel
{
    class ParserExceptions : ArgumentException
    {
        public override string Message { get; }

        protected ParserExceptions()
        {
            Message = "ParserExceptions";
        }
    }

    class InvalidReferenceIndexing : ParserExceptions
    {
        public override string Message { get; }

        public InvalidReferenceIndexing()
        {
            Message = "InvalidReferenceIndexing";
        }
    }

    class DivByZero : ParserExceptions
    {
        public override string Message { get; }

        public DivByZero()
        {
            Message = "DivByZero";
        }
    }

    class InvalidOperation : ParserExceptions
    {
        public override string Message { get; }

        public InvalidOperation()
        {
            Message = "InvalidOperation";
        }
    }

    class InvalidReferenceFormat : ParserExceptions
    {
        public override string Message { get; }

        public InvalidReferenceFormat()
        {
            Message = "InvalidReferenceFormat";
        }
    }

    class RecursiveReference : ParserExceptions
    {
        public override string Message { get; }

        public RecursiveReference()
        {
            Message = "RecursiveReference";
        }
    }
}
