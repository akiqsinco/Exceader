# Exceader
Exceader is a lightweight Excel data reader.

Features:

- Simple and intuitive APIs
- Less dependencies

## Supported formats
This software supports only Excel 2007 or later format (*.xlsx) currently.

## Usage
### Get a cell value
You can get a cell value by the index or Excell's number.

```csharp
using (var book = Book.Open("/path/to/file"))
{
    var sheet = book["sheet name"];

    // by index
    var a1 = sheet[0, 0].Value;
    var b2 = sheet[1][1].Value;

    // by Excel's number
    var c3 = sheet["C3"].Value;
    var d4 = sheet[3]["D"].Value;

    // Get a value as other types
    var @int = sheet["A1"].AsInteger();
    var @float = sheet["A1"].AsFloat();
    var @double = sheet["A1"].AsDouble();
    var dt1 = sheet["A1"].AsDateTime();
    var dt2 = sheet["A1"].AsDateTime("yyyy/MM/dd");
}
```

### Iteration
In Exceader, row and column has infinity length (actually `int.MaxLength`). So you need to specify the range to read for iteration.

```csharp
using (var book = Book.Open("/path/to/file"))
{
    var sheet = book["sheet name"];

    foreach (var row in sheet.Range(0, 16))
    {
        // by column index
        foreach (var cell in row.Range(2, 8))
        {
            var columnIndex = cell.Index;
        }

        // by column name
        foreach (var cell in row.Range("C", "I"))
        {
            var columnName = cell.ColumnName;
        }
    }
}
```

## Installation
You can install via NuGet.

```sh
PM> Install-Package Exceader
```

## Author
- akiqsinco (<akiqsinco@gmail.com>)

## Copyright
Copyright (c) 2019 akiqsinco (<akiqsinco@gmail.com>)

## License
This software is released under the MIT License, see LICENSE.txt.