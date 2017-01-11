using System;
using System.Runtime.Serialization;

namespace CalcEngine
{
    [Serializable()]
    public class CircularReferenceException : Exception
    {
        private string _message = string.Empty;
        public override string Message { get { return string.Format("Circular reference detected {0}", _message); } }

        internal CircularReferenceException()
        {
        }

        public CircularReferenceException(string error) : this(error, null)
        {
        }

        private CircularReferenceException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        public CircularReferenceException(string error, Exception innerException) : base(error, innerException)
        {
            if (!string.IsNullOrEmpty(error))
                _message = error;
        }
    }
}
