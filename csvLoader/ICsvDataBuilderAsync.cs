using System.Threading.Tasks;

namespace excelToolsCore.csvLoader
{
    public interface ICsvDataBuilderAsync<T>
    {
        Task<T> GetTAsync(string[] data);
    }
}