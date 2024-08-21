namespace SimpleExcelExporter.Definitions
{
  using System;

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
  }
}
