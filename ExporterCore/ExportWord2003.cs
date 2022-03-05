using RazorEngineCore;

namespace ExporterCore
{
    public class ExportWord2003<T> : Export<T> where T : class
    {
        private const string Word2003Header = "Word2003Header";
        private const string Word2003Item = "Word2003Item";
        public ExportWord2003()
        {
            var props = properties.Select(it => it.Name).ToArray();
            ExportCollection = Templates.Word2003File;
            ExportHeader = Word2003RazorTemplate.ExportHeaderCompiled.Run(props);
            ExportItem = Word2003RazorTemplate.ExportItemCompiled.Run(props);
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();

            IDictionary<string, string> parts = new Dictionary<string, string>()
            {
                {TType.Name + Word2003Header, ExportHeader},
                {TType.Name + Word2003Item, ExportItem}
            };

            IncludeCompiledTemplate compiledTemplate = razorEngine.Compile(ExportCollection, parts);
            string result = compiledTemplate.Run(modelTemplate);
            return System.Text.Encoding.UTF8.GetBytes(result);
        }

    }

    public class Word2003RazorTemplate
    {
        /// <summary>
        /// excel的每一项
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportItemCompiled;

        /// <summary>
        /// excel的头部
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportHeaderCompiled;

        static Word2003RazorTemplate()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            ExportHeaderCompiled = razorEngine.Compile(Templates.Word2003Header);
            ExportItemCompiled = razorEngine.Compile(Templates.Word2003Item);
        }
    }
}