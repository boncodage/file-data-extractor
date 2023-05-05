using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelExtractor.ExcelExtractor;

public abstract class BaseExcelDataExtractor<T>
{
    protected abstract T ExtraerDatos(ISheet sheet);

    protected internal async Task<T> ExtraerDatosAsync(string path)
    {
        await using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
        {
            var fileExtension = Path.GetExtension(path);
            if (fileExtension == ".xlsx")
            {
                using (var wb = new XSSFWorkbook(fileStream))
                {
                    var sheet = wb.GetSheetAt(0);
                    return await Task.FromResult(ExtraerDatos(sheet));
                }
            } else if(fileExtension == ".xls")
            {
                using (var wb = new HSSFWorkbook(fileStream))
                {
                    var sheet = wb.GetSheetAt(0);
                    return await Task.FromResult(ExtraerDatos(sheet));
                }
            }

            return default;
        }
    }
}