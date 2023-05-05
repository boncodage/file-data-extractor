// See https://aka.ms/new-console-template for more information

using System.Text.Json;
using ExcelExtractor.Archivo;

string path = "archivo.csv";
var archivoEsbExcel = new ArchivoCsvFileDataExtractor();
var result = archivoEsbExcel.ExtraerDatosArchivo(path).ToList();
// utilizar los datos extraídos...
Console.WriteLine(JsonSerializer.Serialize(result));