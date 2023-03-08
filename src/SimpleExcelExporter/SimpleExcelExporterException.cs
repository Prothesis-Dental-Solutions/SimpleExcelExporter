namespace SimpleExcelExporter
{
  using System;
  using System.Runtime.Serialization;

  public class SimpleExcelExporterException : Exception
  {
    // ReSharper disable once UnusedMember.Global
    public SimpleExcelExporterException()
    {
    }

    public SimpleExcelExporterException(
      string message)
  : base(message)
    {
    }

    public SimpleExcelExporterException(
      string message,
      Exception innerException)
      : base(message, innerException)
    {
    }

    protected SimpleExcelExporterException(
      SerializationInfo serializationInfo,
      StreamingContext streamingContext)
      : base(serializationInfo, streamingContext)
    {
    }
  }
}
