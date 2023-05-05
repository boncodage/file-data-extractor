# File Data Extractor Project
Console app for extracting data from XLS, XLSX and CSV files with NPOI and CsvHelper, with SRP (Single Responsability Principle) and decoupling of classes

## Usage
```csharp
using System.Text.Json;
using ExcelExtractor.Archivo;

string path = "archivo.csv";
var extractorArchivo = new ArchivoCsvFileDataExtractor();
var result = extractorArchivo.Extraer(path).ToList();
// utilizar los datos extraídos...
Console.WriteLine(JsonSerializer.Serialize(result));
