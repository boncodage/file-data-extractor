namespace ExcelExtractor;

public static class Enums
{
    /// <summary>
    /// Obtiene todos los valores de un Enum como un Array.
    /// </summary>
    /// <typeparam name="TEnum"></typeparam>
    /// <returns></returns>
    public static Array AsArray<TEnum>() where TEnum : struct, Enum =>
        typeof(TEnum).GetEnumValuesAsUnderlyingType();
}