using System.Reflection;

namespace ExporterCore
{
    public abstract class Export<T>
    {
        protected PropertyInfo[] properties;
        protected Type TType;

        protected Export()
        {
            TType= typeof (T);
            properties = TType.GetProperties();
            
        }
        public string ExportCollection { get; set; }
        public string ExportItem { get; set; }
        public string ExportHeader { get; set; }

        public abstract byte[] ExportResult(List<T> data,params KeyValuePair<string, object>[] additionalData);

    }
}