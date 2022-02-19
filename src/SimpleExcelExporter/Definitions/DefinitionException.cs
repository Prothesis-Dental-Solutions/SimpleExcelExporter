namespace SimpleExcelExporter.Definitions
{
  using System;
  using System.Runtime.Serialization;

  [Serializable]
  public class DefinitionException : Exception
  {
    public DefinitionException(string message)
      : base(message)
    {
    }

    public DefinitionException(string message, Exception innerException)
      : base(message, innerException)
    {
    }

    public DefinitionException()
    {
    }

    protected DefinitionException(SerializationInfo serializationInfo, StreamingContext streamingContext)
      : base(serializationInfo, streamingContext)
    {
    }
  }
}
