// See https://aka.ms/new-console-template for more information

using System.Text.Json;
using ExcelExtractor.ArchivoEsb;

string path = "archivo.csv";
var archivoEsbExcel = new ArchivoEsbCsvDataExtractor();
var result = archivoEsbExcel.ExtraerDatosDesdeArchivo(path).ToList();
// utilizar los datos extraídos...
Console.WriteLine(JsonSerializer.Serialize(result));