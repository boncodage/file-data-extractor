using NPOI.SS.UserModel;

namespace ExcelExtractor.Archivo;

public class ArchivoExcelFileDataExtractor : BaseFileDataExtractor<IEnumerable<IFila<ArchivoFilaModel>>>
{
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
