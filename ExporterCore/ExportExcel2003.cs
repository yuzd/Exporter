using RazorEngineCore;

namespace ExporterCore
{
    public class ExportExcel2003<T> : Export<T> where T : class
    {
        private const string Excel2003Header = "Excel2003Header";
        private const string Excel2003Item = "Excel2003Item";
        public ExportExcel2003()
        {
            var props = properties.Select(it => it.Name).ToArray();
            ExportCollection = Templates.Excel2003File;
            ExportHeader = Excel2003RazorTemplate.ExportHeaderCompiled.Run(props);
            ExportItem = Excel2003RazorTemplate.ExportItemCompiled.Run(props);
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {

            var modelTemplate = new ModelTemplate<T>(data);
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();

            IDictionary<string, string> parts = new Dictionary<string, string>()
            {
                {TType.Name + Excel2003Header, ExportHeader},
                {TType.Name + Excel2003Item, ExportItem}
            };

            IncludeCompiledTemplate compiledTemplate = razorEngine.Compile(ExportCollection, parts);
            string result = compiledTemplate.Run(modelTemplate);
            return System.Text.Encoding.UTF8.GetBytes(result);

        }

    }

    public class Excel2003RazorTemplate
    {
        /// <summary>
        /// excel的每一项
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportItemCompiled;

        /// <summary>
        /// excel的头部
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportHeaderCompiled;

        static Excel2003RazorTemplate()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            ExportHeaderCompiled = razorEngine.Compile(Templates.Excel2003Header);
            ExportItemCompiled = razorEngine.Compile(Templates.Excel2003Item);
        }
    }
}