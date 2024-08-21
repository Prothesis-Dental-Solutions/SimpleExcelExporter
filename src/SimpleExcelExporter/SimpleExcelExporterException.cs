namespace SimpleExcelExporter
{
  using System;

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
  }
}
