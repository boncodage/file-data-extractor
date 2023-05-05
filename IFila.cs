namespace ExcelExtractor;

public interface IFila<T>
{
    public bool HasErrors => Errors?.Any() ?? false;
    public IList<string> Errors { get; set; }
    public T Data { get; set; }
}