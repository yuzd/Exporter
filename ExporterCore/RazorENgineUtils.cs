using RazorEngineCore;

namespace ExporterCore
{
    public class IncludeTemplateBase : RazorEngineTemplateBase
    {
        public Func<string, object?, string> IncludeCallback { get; set; }
        public Func<string> RenderBodyCallback { get; set; }
        public string Layout { get; set; }

        public string Include(string key, object? model = null)
        {
            return this.IncludeCallback(key, model);
        }

        public string RenderBody()
        {
            return this.RenderBodyCallback();
        }
    }



    public class IncludeCompiledTemplate
    {
        private readonly IRazorEngineCompiledTemplate<IncludeTemplateBase> compiledTemplate;
        private readonly Dictionary<string, IRazorEngineCompiledTemplate<IncludeTemplateBase>> compiledParts;

        public IncludeCompiledTemplate(IRazorEngineCompiledTemplate<IncludeTemplateBase> compiledTemplate, Dictionary<string, IRazorEngineCompiledTemplate<IncludeTemplateBase>> compiledParts)
        {
            this.compiledTemplate = compiledTemplate;
            this.compiledParts = compiledParts;
        }

        public string Run(object model)
        {
            return this.Run(this.compiledTemplate, model);
        }

        public string Run(IRazorEngineCompiledTemplate<IncludeTemplateBase> template, object model)
        {
            IncludeTemplateBase? templateReference = null;

            string result = template.Run(instance =>
            {
                if (!(model is AnonymousTypeWrapper))
                {
                    model = new AnonymousTypeWrapper(model);
                }

                instance.Model = model;
                instance.IncludeCallback = (key, includeModel) => this.Run(this.compiledParts[key], includeModel!);

                templateReference = instance;
            });

            if (templateReference?.Layout == null)
            {
                return result;
            }

            return this.compiledParts[templateReference.Layout].Run(instance =>
            {
                if (!(model is AnonymousTypeWrapper))
                {
                    model = new AnonymousTypeWrapper(model);
                }

                instance.Model = model;
                instance.IncludeCallback = (key, includeModel) => this.Run(this.compiledParts[key], includeModel!);
                instance.RenderBodyCallback = () => result;
            });
        }
       
    }
}
