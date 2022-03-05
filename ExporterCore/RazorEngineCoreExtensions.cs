using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RazorEngineCore;

namespace ExporterCore
{
    public static class RazorEngineCoreExtensions
    {
        public static IncludeCompiledTemplate Compile(this IRazorEngine razorEngine, string template, IDictionary<string, string> parts)
        {
            return new IncludeCompiledTemplate(
                razorEngine.Compile<IncludeTemplateBase>(template),
                parts.ToDictionary(
                    k => k.Key,
                    v => razorEngine.Compile<IncludeTemplateBase>(v.Value)));
        }

       
    }

}
