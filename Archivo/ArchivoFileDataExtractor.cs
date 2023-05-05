using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using NPOI.SS.UserModel;

namespace ExcelExtractor.Archivo;

public class ArchivoFileDataExtractor : BaseFileDataExtractor<IEnumerable<IFila<ArchivoFilaModel>>>
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
    protected override IEnumerable<IFila<ArchivoFilaModel>> ExtraerDatosExcel(ISheet sheet)
    {
        var rows = sheet.GetRowEnumerator();
        while (rows.MoveNext())
        {
            var row = (IRow)rows.Current;
            if (row != null)
            {
                var datosFila = new ArchivoFilaModel();
                var errores = new List<string>();
                
                row.Cells.ForEach(celda =>
                {
                    var columna = (ColumnasArchivo) celda.ColumnIndex;
                    try
                    {
                        switch (columna)
                        {
                            case ColumnasArchivo.Servicio:
                                datosFila.Servicio = row.GetCell((int) ColumnasArchivo.Servicio)?.StringCellValue;
                                break;
                            case ColumnasArchivo.Metodo:
                                datosFila.Metodo = row.GetCell((int) ColumnasArchivo.Metodo)?.StringCellValue;
                                break;
                            case ColumnasArchivo.Url:
                                datosFila.Url = row.GetCell((int) ColumnasArchivo.Url)?.StringCellValue;
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                        errores.Add($"Error de lectura: {e.Message}, Celda: {celda.Address}");
                    }
                });
                
                yield return new Fila<ArchivoFilaModel>()
                {
                    Data = datosFila,
                    Errors = errores
                };
            }
        }
    }
}
