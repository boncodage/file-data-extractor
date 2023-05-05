// See https://aka.ms/new-console-template for more information

using System.Text.Json;
using ExcelExtractor.Archivo;

string path = "archivo.csv";
var extractorArchivo = new ArchivoCsvFileDataExtractor();
var result = extractorArchivo.Extraer(path).ToList();
// utilizar los datos extraídos...
Console.WriteLine(JsonSerializer.Serialize(result));
