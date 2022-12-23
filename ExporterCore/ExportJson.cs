using Newtonsoft.Json;

namespace ExporterCore
{
    public class ExportJson<T> : Export<T> where T : class
    {
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            string result = JsonConvert.SerializeObject(data);
            return System.Text.Encoding.UTF8.GetBytes(result);
        }
    }
}