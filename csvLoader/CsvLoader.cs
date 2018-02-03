using System.Collections.Generic;
using System.Threading.Tasks;

namespace excelToolsCore.csvLoader
{
    public class CsvLoader
    {
        private int startingRow = 1;

        public List<T> GetData<T, U>(string filePath) where U : ICsvDataBuilder<T>, new()
        {
            List<T> buildedData = new List<T>();
            string[][] rawData = utilities.Helpers.ReadeCsv(filePath);

            for (int i = startingRow; i < rawData.Length; i++)
            {
                T t = new U().Build(rawData[i]);
                buildedData.Add(t);
            }

            return buildedData;
        }

        public List<T> GetData<T, U>(string[][] data) where U : ICsvDataBuilder<T>, new()
        {
            List<T> buildedData = new List<T>();
            string[][] rawData = data;

            for (int i = startingRow; i < rawData.Length; i++)
            {
                T t = new U().Build(rawData[i]);
                buildedData.Add(t);
            }

            return buildedData;
        }

        public async Task<List<T>> GetDataAsync<T, U>(string filePath) where U : ICsvDataBuilderAsync<T>, new()
        {
            List<T> buildedData = new List<T>();
            string[][] rawData = utilities.Helpers.ReadeCsv(filePath);

            for (int i = startingRow; i < rawData.Length; i++)
            {
                T t = await new U().GetTAsync(rawData[i]).ConfigureAwait(false);
                buildedData.Add(t);
            }

            return buildedData;
        }

        public async Task<List<T>> GetDataAsync<T, U>(string[][] data) where U : ICsvDataBuilderAsync<T>, new()
        {
            List<T> buildedData = new List<T>();
            string[][] rawData = data;

            for (int i = startingRow; i < rawData.Length; i++)
            {
                T t = await new U().GetTAsync(rawData[i]).ConfigureAwait(false);
                buildedData.Add(t);
            }

            return buildedData;
        }

        public void SetStartingRow(int row) => startingRow = row;
    }
}