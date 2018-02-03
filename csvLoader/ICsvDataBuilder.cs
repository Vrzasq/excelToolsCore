namespace excelToolsCore.csvLoader
{
    public interface ICsvDataBuilder<T>
    {
        T Build(string[] data);
    }
}