﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExporterCore;
using ExporterObjects;
using Microsoft.SqlServer.Server;
using RazorEngine;
using RazorEngine.Templating;

namespace ExportImplementation
{
    public class ExportWord2003<T> : Export<T>
        where T : class
    {
        public ExportWord2003()
        {
            ExportCollection = Templates.Word2003File;
            string template = Templates.Word2003Header;
            var props = properties.Select(it => it.Name).ToArray();
            ExportHeader = Engine.Razor.RunCompile(template,TType.Name + "Word2003HeaderInterpreter",  typeof (string[]),props);
            template = Templates.Word2003Item;
            ExportItem = Engine.Razor.RunCompile(template, TType.Name + "Word2003ItemInterpreter", typeof(string[]), props);

        }

        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            
            var service = Engine.Razor;
            service.AddTemplate(TType.Name + "Word2003Collection",ExportCollection);
            service.AddTemplate(TType.Name + "Word2003Header", ExportHeader);
            service.AddTemplate(TType.Name + "Word2003Item", ExportItem);
            service.Compile(TType.Name + "Word2003Collection",typeof(ModelTemplate<T>));
            service.Compile(TType.Name + "Word2003Header");
            service.Compile(TType.Name + "Word2003Item",typeof(T));
            var result = service.Run(TType.Name + "Word2003Collection", typeof(ModelTemplate<T>), modelTemplate, additionalData.ToDynamicViewBag());
            return System.Text.Encoding.UTF8.GetBytes(result);


        }

    }
}