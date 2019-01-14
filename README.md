# Exceader
Exceader is a lightweight Excel data reader.  

Features:

- Simple API
- Less dependencies
- Support Office 2007 or later

## Warning
This software is still BETA quality. The APIs will be likely to change.

## Installation
WIP

## Usage
```csharp
using (var book = Book.Open("/path/to/excel file"))
{
    var sheet = book["sheet name"];

    // Get a cell value
    var a1 = sheet["A1"].Value;
    var b2 = sheet[1, 1].Value;
    var c3 = sheet[2][2].Value;

    // Get a value as other types
    var intVal = sheet["AC12"].AsInteger();
    var doubleVal = sheet["AC12"].AsDouble();
    var dateTimeVal = sheet["AC12"].AsDateTime();
    var dateTimeVal = sheet["AC12"].AsDateTime("yyyy/MM/dd");

    // Iteration
    foreach (var row in sheet.Range(0, 16))
    {
        foreach (var cell in row.Range(2, 8))
        {
            var val = cell.Value;
        }
    }
}
```

## Author
- akiqsinco (<akiqsinco@gmail.com>)

## Copyright
Copyright (c) 2019 akiqsinco (<akiqsinco@gmail.com>)

## License
This software is released under the MIT License, see LICENSE.txt.