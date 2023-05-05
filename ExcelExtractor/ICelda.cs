namespace ExcelExtractor.ExcelExtractor;

public interface ICelda<TResult>
{
    public string Id { get; set; }
    public bool Error { get; set; }
    public string Mensaje { get; set; }
    public TResult Valor { get; set; }
    
}