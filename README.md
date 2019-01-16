# Exceader
Exceader is a lightweight Excel data reader.

Features:

- Simple and intuitive APIs
- Less dependencies

## Supported formats
This software supports only Excel 2007 or later format (*.xlsx) currently.

## Usage
```csharp
using (var book = Book.Open("/path/to/file"))
{
    var sheet = book["sheet name"];
    var row = sheet[2];

    // Get a cell value
    var a1 = sheet["A1"].Value;
    var b2 = sheet[1, 1].Value;
    var c3 = row["C"].Value;
    var d3 = row[3].Value;

    // Get a value as other types
    var @double = sheet["AC12"].AsDouble();
    var dateTime = sheet["AC12"].AsDateTime();

    // Iteration
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
WIP

## Author
- akiqsinco (<akiqsinco@gmail.com>)

## Copyright
Copyright (c) 2019 akiqsinco (<akiqsinco@gmail.com>)

## License
This software is released under the MIT License, see LICENSE.txt.