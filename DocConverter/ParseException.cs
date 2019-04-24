using System;
using System.Runtime.Serialization;

namespace DocConverter
{
    [Serializable]
    public class ParseException : Exception, ISerializable
    {
        public ParseException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
