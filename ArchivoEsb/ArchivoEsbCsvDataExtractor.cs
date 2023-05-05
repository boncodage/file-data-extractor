using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using ExcelExtractor.ExcelExtractor;

namespace ExcelExtractor.ArchivoEsb;

public class ArchivoEsbCsvDataExtractor : BaseCsvDataExtractor<IEnumerable<IFila<EsbArchivoFilaModel>>>
{
    public ArchivoEsbCsvDataExtractor(): base(delimitador: ';')
    {
        
    }
    protected override IEnumerable<IFila<EsbArchivoFilaModel>> ExtraerDatos(string path)
    {
        using var reader = new StreamReader(path, new FileStreamOptions(){ Access = FileAccess.Read, Mode = FileMode.Open});
        using var csvReader = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture){ Delimiter = ";"});
        
        while (csvReader.Read())
        {
            var datosFila = new EsbArchivoFilaModel();
            var errores = new List<string>();
            var columnas = Enums.AsArray<ColumnasArchivoEsb>();
            for (int iColumna = 0; iColumna < columnas.Length; iColumna++)
            {
                var columna = (ColumnasArchivoEsb) iColumna;
                try
                {
                    switch (columna)
                    {
                        case ColumnasArchivoEsb.Servicio:
                            datosFila.Servicio = csvReader.GetField<string>(iColumna);
                            break;
                        case ColumnasArchivoEsb.Metodo:
                            datosFila.Metodo = csvReader.GetField<DateTime?>(iColumna);
                            break;
                        case ColumnasArchivoEsb.Url:
                            datosFila.Url = csvReader.GetField<string>(iColumna);
                            break;
                    }
                }
                catch (TypeConverterException e)
                {
                    errores.Add($"Error de lectura. ['{e.Text}' -> {e.TypeConverter} ], Fila: {csvReader.CurrentIndex + 1}, Columna: {iColumna + 1}");
                }
            }

            yield return new Fila<EsbArchivoFilaModel>()
            {
                Data = datosFila,
                Errors = errores
            };
        }
    }

}
