using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using RazorEngineCore;

namespace ExporterCore
{
    public class ExportWord<T> : Export<T> where T : class
    {
        private const string Word2007Header = "Word2007Header";
        private const string Word2007Item = "Word2007Item";

        public ExportWord()
        {
            var props = properties.Select(it => it.Name).ToArray();
            ExportCollection = Templates.Word2007File;
            ExportHeader = Word2007RazorTemplate.ExportHeaderCompiled.Run(props);
            ExportItem = Word2007RazorTemplate.ExportItemCompiled.Run(props);

        }

        public string ExportResultStringPart(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();

            IDictionary<string, string> parts = new Dictionary<string, string>()
            {
                {TType.Name + Word2007Header, ExportHeader},
                {TType.Name + Word2007Item, ExportItem}
            };

            IncludeCompiledTemplate compiledTemplate = razorEngine.Compile(ExportCollection, parts);
            string result = compiledTemplate.Run(modelTemplate);
            return result;
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var result = ExportResultStringPart(data, additionalData);
            return CreateWord2007(result);


        }
        private byte[] CreateWord2007(string Text)
        {
            using var ms = new MemoryStream();
            using var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document);
            // Set the content of the document so that Word can open it.
            var mainPart = wordDoc.AddMainDocumentPart();
            WriteToPart(mainPart, Text);
            wordDoc.Close();
            return ms.ToArray();
        }
        /// <summary>
        /// TODO: move into utilities
        /// </summary>
        /// <param name="oxp"></param>
        /// <param name="Text"></param>

        internal void WriteToPart(OpenXmlPart oxp, string Text)
        {
            using Stream stream = oxp.GetStream();
            byte[] buf = (new UTF8Encoding()).GetBytes(Text);
            stream.Write(buf, 0, buf.Length);
        }
    }
    public class Word2007RazorTemplate
    {
        /// <summary>
        /// excel的每一项
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportItemCompiled;

        /// <summary>
        /// excel的头部
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportHeaderCompiled;

        static Word2007RazorTemplate()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            ExportHeaderCompiled = razorEngine.Compile(Templates.Word2007Header);
            ExportItemCompiled = razorEngine.Compile(Templates.Word2007Item);
        }
    }
}