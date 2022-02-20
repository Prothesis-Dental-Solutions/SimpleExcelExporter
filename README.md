# SimpleExcelExporter

[![NuGet version (SimpleExcelExporter)](https://img.shields.io/nuget/v/SimpleExcelExporter.svg?style=flat-square)](https://www.nuget.org/packages/SimpleexcelExporter/)
[![.NET](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/actions/workflows/dotnet.yml/badge.svg)](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/actions/workflows/dotnet.yml)

This C# library helps export data to Excel .xlsx file.

Internally it uses the [Microsoft DocumentFormat.OpenXml library](https://github.com/OfficeDev/Open-XML-SDK). Unlike DocumentFormat.OpenXml, this library doesn't aim to produce any possible kind of .xlsx file but focus on the use case where a user of your application wants to export data to an Excel file. Of course you can use DocumentFormat.OpenXml directly but we believe this library is simpler and is less error prone for that particular use case.

## How to use?
1. Let's say you have a Player class in your code.
``` C#
public class Player
{
  public DateTime? DateOfBirth { get; set; }

  public string? PlayerName { get; set; }

  public int? NumberOfVictory { get; set; }

  public bool? IsActiveFlag { get; set; }

  public double? Size { get; set; }

  public decimal? Salary { get; set; }
}
```
and a team class
``` C#
public class Team
{
  private ICollection<Player>? _players;

  public ICollection<Player> Players
  {
    get => _players ??= new HashSet<Player>();
  }
}
```

2. Use SimpleExcelExporter to generate an Excel file:
The following snippet comes from [SimpleExcelEporsterExaple](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporterExample/blob/b7a3b184892f83370b8d937dffd824c7f82b0b06/Program.cs#L1-L42)
``` C#
// Instanciate the objects to export
var team = new Team
{
  Players =
{
  new Player
  {
    PlayerName = "Alexandre",
    Size = 1.93d,
    DateOfBirth = new DateTime(1974, 02, 01),
    IsActiveFlag = true,
    NumberOfVictory = 45,
    Salary = 2000.50m,
  },
  new Player
  {
    PlayerName = "Elina",
    Size = 1.72d,
    DateOfBirth = new DateTime(1990, 10, 13),
    IsActiveFlag = true,
    NumberOfVictory = 52,
    Salary = 2141.25m,
  }
},
};

// Create a temp directory
var n = DateTime.Now;
var tempDi = new DirectoryInfo($"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");
tempDi.Create();

// Export team to an excel file
using var memoryStream = new MemoryStream();
using var streamWriter = new StreamWriter(memoryStream);
SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, team);
spreadsheetWriter.Write();
using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "Team.xlsx"), FileMode.Create, FileAccess.Write);
memoryStream.WriteTo(file);
```

## How to configure export ?
1. If you want to configure the column names and/or the sheet name, create a resource file.
In the following code I assume you have a resource file 

2. Anotate your classes
The following snippet comes from [SimpleExcelEporsterExaple](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporterExample/blob/b7a3b184892f83370b8d937dffd824c7f82b0b06/Team.cs#L1-L16)
``` C#
public class Team
{
  private ICollection<Player>? _players;

  [SheetName(typeof(TeamRes), "SheetName")]
  [EmptyResultMessage(typeof(TeamRes), "EmptyResultMessage")]
  public ICollection<Player> Players
  {
    get => _players ??= new HashSet<Player>();
  }
}
```

The following snippet comes from [SimpleExcelEporsterExaple](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporterExample/blob/b7a3b184892f83370b8d937dffd824c7f82b0b06/Player.cs#L6-L37)
``` C#
public class Player
{
  [CellDefinition(CellDataType.Date)]
  [Header(typeof(TeamRes), "DateOfBirthColumnName")]
  [Index(2)]
  public DateTime? DateOfBirth { get; set; }

  [CellDefinition(CellDataType.String)]
  [Header(typeof(TeamRes), "PlayerNameColumnName")]
  [Index(1)]
  public string? PlayerName { get; set; }

  [CellDefinition(CellDataType.Number)]
  [Header(typeof(TeamRes), "NumberOfVictoryColumnName")]
  [Index(3)]
  public int? NumberOfVictory { get; set; }

  [CellDefinition(CellDataType.Boolean)]
  [Header(typeof(TeamRes), "IsActiveFlagColumnName")]
  [Index(4)]
  public bool? IsActiveFlag { get; set; }

  [CellDefinition(CellDataType.Boolean)]
  [Header(typeof(TeamRes), "SizeColumnName")]
  [Index(5)]
  public double? Size { get; set; }

  [CellDefinition(CellDataType.Number)]
  [Header(typeof(TeamRes), "SalaryColumnName")]
  [Index(5)]
  public decimal? Salary { get; set; }
}
```

## What if I don't want to annotate my class with your annotations?
You can still use SimpleExcelExporter but you'll have a bit more code to write. Have a look at [this example](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/blob/dda3b06649b6ec9e4126c0f5af743c931c048595/src/ConsoleApp/Program.cs#L211-L274)
