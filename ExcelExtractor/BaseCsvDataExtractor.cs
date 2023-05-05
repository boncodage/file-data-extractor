using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;

namespace ExcelExtractor.ExcelExtractor;

public abstract class BaseCsvDataExtractor<T>
{
    private readonly char _delimitador;

    protected BaseCsvDataExtractor(char delimitador)
    {
        _delimitador = delimitador;
    }
    protected abstract T ExtraerDatos(string path);

    protected internal T ExtraerDatosDesdeArchivo(string path)
    {
        return ExtraerDatos(path);
    }
}