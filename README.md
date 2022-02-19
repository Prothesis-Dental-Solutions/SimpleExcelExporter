# SimpleExcelExporter

This C# library is intended to help export data to Excel .xlsx file.

Internally it uses the [Microsoft DocumentFormat.OpenXml library](https://github.com/OfficeDev/Open-XML-SDK). Unlike DocumentFormat.OpenXml, this library doesn't aim to produce any possible kind of .xlsx file but focus on the use case where a user of your application wants to export data to an Excel file. Of course you can use DocumentFormat.OpenXml directly but we believe this library is simpler and is less error prone for that particular use case.

## How to use?
1. Annotate your classes. Let's say you have a Player class in your code. Add annotations to guide SimpleExcelExporter how to export your data.
``` C#
public class Player
{
  [CellDefinition(CellDataType.Date)]
  [Header(typeof(PlayerRes), "DateOfBirthColumnName")]
  [Index(2)]
  public DateTime? DateOfBirth { get; set; }

  [CellDefinition(CellDataType.String)]
  [Header(typeof(PlayerRes), "PlayerNameColumnName")]
  [Index(1)]
  public string? PlayerName { get; set; }

  [CellDefinition(CellDataType.Number)]
  [Header(typeof(PlayerRes), "NumberOfVictoryColumnName")]
  [Index(3)]
  public int? NumberOfVictory { get; set; }

  [CellDefinition(CellDataType.Boolean)]
  [Header(typeof(PlayerRes), "IsActiveFlagColumnName")]
  [Index(4)]
  public bool? IsActiveFlag { get; set; }

  [CellDefinition(CellDataType.Boolean)]
  [Header(typeof(PlayerRes), "SizeColumnName")]
  [Index(5)]
  public double? Size { get; set; }

  [CellDefinition(CellDataType.Number)]
  [Header(typeof(PlayerRes), "SalaryColumnName")]
  [Index(5)]
  public decimal? Salary { get; set; }
}
```
2. Use SimpleExcelExporter to generate an Excel file:
``` C#
using var memoryStream = new MemoryStream();
using var streamWriter = new StreamWriter(memoryStream);
var team = new Team
{
Players =
{
  new Player
  {
	PlayerCode = "01",
	PlayerName = "Alexandre",
	Size = 1.93d,
	DateOfBirth = new DateTime(1974, 02, 01),
	IsActiveFlag = true,
	NumberOfVictory = 45,
	Salary = 2000.50m,
  },
  new Player
  {
	PlayerCode = "02",
	PlayerName = "Elina",
	Size = 1.72d,
	DateOfBirth = new DateTime(1990, 10, 13),
	IsActiveFlag = true,
	NumberOfVictory = 52,
	Salary = 2141.25m,
  }
},
};

SpreadsheetWriter spreadsheetWriter = new SpreadsheetWriter(streamWriter.BaseStream, team);
spreadsheetWriter.Write();

using FileStream file = new FileStream(Path.Combine(tempDi.FullName, "TestWithData2.xlsx"), FileMode.Create, FileAccess.Write);
memoryStream.WriteTo(file);
```

## What if I don't want to annotate my class with your annotations?
You can still use SimpleExcelExporter but you'll have a bit more code to write. Have a look at [this example](https://github.com/Prothesis-Dental-Solutions/SimpleExcelExporter/blob/dda3b06649b6ec9e4126c0f5af743c931c048595/src/ConsoleApp/Program.cs#L211-L274)
