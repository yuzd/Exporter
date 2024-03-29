﻿using RazorEngineCore;

namespace ExporterCore
{
    public class ExportHtml<T> : Export<T> where T : class
    {
        private const string HtmlHeader = "HtmlHeader";
        private const string HtmlItem = "HtmlItem";
        public ExportHtml()
        {
            var props = properties.Select(it => it.Name).ToArray();
            ExportCollection = Templates.HtmlFile;
            ExportHeader = HtmlRazorTemplate.ExportHeaderCompiled.Run(props);
            ExportItem = HtmlRazorTemplate.ExportItemCompiled.Run(props);
        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();

            IDictionary<string, string> parts = new Dictionary<string, string>()
            {
                {TType.Name + HtmlHeader, ExportHeader},
                {TType.Name + HtmlItem, ExportItem}
            };

            IncludeCompiledTemplate compiledTemplate = razorEngine.Compile(ExportCollection, parts);
            string result = compiledTemplate.Run(modelTemplate);
            return System.Text.Encoding.UTF8.GetBytes(result);

        }

    }


    public class HtmlRazorTemplate
    {
        /// <summary>
        /// excel的每一项
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportItemCompiled;

        /// <summary>
        /// excel的头部
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportHeaderCompiled;

        static HtmlRazorTemplate()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            ExportHeaderCompiled = razorEngine.Compile(Templates.HtmlHeader);
            ExportItemCompiled = razorEngine.Compile(Templates.HtmlItem);
        }
    }
}