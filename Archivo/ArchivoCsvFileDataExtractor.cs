using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;

namespace ExcelExtractor.Archivo;

public class ArchivoCsvFileDataExtractor : BaseFileDataExtractor<IEnumerable<IFila<ArchivoFilaModel>>>
{
    protected override IEnumerable<IFila<ArchivoFilaModel>> ExtraerDatosCsv(string path)
    {
        using var reader = new StreamReader(path, new FileStreamOptions() {Access = FileAccess.Read, Mode = FileMode.Open});
        using var csvReader = new CsvReader(reader,
            new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ";"
            });

        while (csvReader.Read())
        {
            var datosFila = new ArchivoFilaModel();
            var errores = new List<string>();
            var columnas = Enums.AsArray<ColumnasArchivo>();
            for (int iColumna = 0; iColumna < columnas.Length; iColumna++)
            {
                var columna = (ColumnasArchivo) iColumna;
                try
                {
                    switch (columna)
                    {
                        case ColumnasArchivo.Servicio:
                            datosFila.Servicio = csvReader.GetField<string>(iColumna);
                            break;
                        case ColumnasArchivo.Metodo:
                            datosFila.Metodo = csvReader.GetField<string>(iColumna);
                            break;
                        case ColumnasArchivo.Url:
                            datosFila.Url = csvReader.GetField<string>(iColumna);
                            break;
                    }
                }
                catch (TypeConverterException e)
                {
                    errores.Add($"Error de lectura. ['{e.Text}' -> {e.TypeConverter} ], Fila: {csvReader.CurrentIndex + 1}, Columna: {iColumna + 1}");
                }
            }

            yield return new Fila<ArchivoFilaModel>()
            {
                Data = datosFila,
                Errors = errores
            };
        }
    }
}