using System.Threading.Tasks;

namespace excelToolsCore.csvLoader
{
    public interface ICsvDataBuilderAsync<T>
    {
        Task BuildAsync(string[] data);
        Task<T> GetTAsync();
    }
}