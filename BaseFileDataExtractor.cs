using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExtractor;

public abstract class BaseFileDataExtractor<T>
{
    // ReSharper disable once UnusedParameter.Global
    protected virtual T ExtraerDatosExcel(ISheet sheet)
    {
        throw new NotImplementedException($"No existe un método de extracción de datos desde Excel para el archivo indicado.");
    }

    protected virtual T ExtraerDatosCsv(string path)
    {
        throw new NotImplementedException($"No existe un método de extracción de datos para el archivo: {Path.GetFileName(path)}");
    }

    protected internal T Extraer(string path)
    {
        var fileExtension = Path.GetExtension(path).ToLower();
        switch (fileExtension)
        {
            case ".xls":
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    using (var wb = new HSSFWorkbook(fileStream))
                    {
                        var sheet = wb.GetSheetAt(0);
                        return ExtraerDatosExcel(sheet);
                    }
                }
            case ".xlsx":
                using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    using (var wb = new XSSFWorkbook(fileStream))
                    {
                        var sheet = wb.GetSheetAt(0);
                        return ExtraerDatosExcel(sheet);
                    }
                }
            case ".csv":
                return ExtraerDatosCsv(path);
            default:
                throw new NotImplementedException("No hay ningún método de extracción asociado a este tipo de archivo.");
        }
    }
}
