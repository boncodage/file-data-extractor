namespace ExcelExtractor.ExcelExtractor;

public class Celda<TResult> : ICelda<TResult>
{
    public TResult Valor { get; set; }

    public string Id { get; set; }
    public bool Error { get; set; }
    public string Mensaje { get; set; }

}