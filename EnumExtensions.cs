namespace ExcelExtractor;

public static class Enums
{
    public static Array AsArray<TEnum>() where TEnum : struct, Enum =>
        typeof(TEnum).GetEnumValuesAsUnderlyingType();
}