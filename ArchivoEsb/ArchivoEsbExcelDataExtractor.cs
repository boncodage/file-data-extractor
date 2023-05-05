using ExcelExtractor.ExcelExtractor;
using NPOI.SS.UserModel;

namespace ExcelExtractor.ArchivoEsb;

public class ArchivoEsbExcelDataExtractor : BaseExcelDataExtractor<IEnumerable<IFila<EsbArchivoFilaModel>>>
{
    protected override IEnumerable<IFila<EsbArchivoFilaModel>> ExtraerDatos(ISheet sheet)
    {
        var rows = sheet.GetRowEnumerator();
        while (rows.MoveNext())
        {
            var row = (IRow)rows.Current;
            if (row != null)
            {
                var datosFila = new EsbArchivoFilaModel();
                var errores = new List<string>();
                
                row.Cells.ForEach(celda =>
                {
                    var columna = (ColumnasArchivoEsb) celda.ColumnIndex;
                    try
                    {
                        switch (columna)
                        {
                            case ColumnasArchivoEsb.Servicio:
                                datosFila.Servicio = row.GetCell((int) ColumnasArchivoEsb.Servicio)?.StringCellValue;
                                break;
                            case ColumnasArchivoEsb.Metodo:
                                datosFila.Metodo = row.GetCell((int) ColumnasArchivoEsb.Metodo)?.DateCellValue;
                                break;
                            case ColumnasArchivoEsb.Url:
                                datosFila.Url = row.GetCell((int) ColumnasArchivoEsb.Url)?.StringCellValue;
                                break;
                        }
                    }
                    catch (Exception e)
                    {
                        errores.Add($"Error de lectura: {e.Message}, Celda: {celda.Address}");
                    }
                });
                
                yield return new Fila<EsbArchivoFilaModel>()
                {
                    Data = datosFila,
                    Errors = errores
                };
            }
        }
    }
}
