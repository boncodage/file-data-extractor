namespace ExcelExtractor;

public class Fila<T> : IFila<T> 
{
    public IList<string> Errors { get; set; } = new List<string>();
    public T Data { get; set; } 
}