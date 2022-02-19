namespace ConsoleApp
{
  using System;

  public class RandomDateTime
  {
    private readonly DateTime _start;
    private readonly Random _gen;
    private readonly int _range;

    public RandomDateTime()
    {
      _start = new DateTime(1970, 1, 1);
      _gen = new Random();
      _range = (DateTime.Today - _start).Days;
    }

    public DateTime Next()
    {
      return _start.AddDays(_gen.Next(_range)).AddHours(_gen.Next(0, 24)).AddMinutes(_gen.Next(0, 60)).AddSeconds(_gen.Next(0, 60));
    }
  }
}
