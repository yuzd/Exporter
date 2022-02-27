﻿using System;
using System.Collections;
using System.Data;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using ExporterObjects;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RazorEngine;
using RazorEngine.Templating;
using Xamasoft.JsonClassGenerator;
using Xamasoft.JsonClassGenerator.CodeWriters;
using Encoding = System.Text.Encoding;

namespace ExportImplementation
{
    public static class ExportFactory
    {
        static ExportFactory()
        {
            NatashaInitializer.InitializeAndPreheating().GetAwaiter().GetResult();
        }
        public static byte[] ExportDataWithType(IEnumerable data, ExportToFormat exportFormat, Type type,
            params KeyValuePair<string, object>[] additionalData)

        {
            var exportType = typeof(Export<>).MakeGenericType(type);

            switch (exportFormat)
            {
                case ExportToFormat.Word2003XML:

                    exportType = typeof(ExportWord2003<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Excel2003XML:
                    exportType = typeof(ExportExcel2003<>).MakeGenericType(type);
                    break;
                case ExportToFormat.HTML:
                    exportType = typeof(ExportHtml<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Word2007:
                    exportType = typeof(ExportWord2007<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Excel2007:
                    exportType = typeof(ExportExcel2007<>).MakeGenericType(type);
                    break;
                default:
                    //throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
                    throw new ArgumentOutOfRangeException("exportFormat", exportFormat, null);
            }

            dynamic export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export.ExportResult(data1, additionalData);

        }

        public static byte[] ExportData<T>(List<T> data, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
            where T : class
        {
            Export<T> export;
            switch (exportFormat)
            {
                case ExportToFormat.Word2003XML:
                    export = new ExportWord2003<T>();
                    break;
                case ExportToFormat.Excel2003XML:
                    export = new ExportExcel2003<T>();
                    break;
                case ExportToFormat.HTML:
                    export = new ExportHtml<T>();
                    break;
                case ExportToFormat.Word2007:
                    export = new ExportWord2007<T>();
                    break;
                case ExportToFormat.Excel2007:
                    export = new ExportExcel2007<T>();
                    break;
                default:
                    //throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
                    throw new ArgumentOutOfRangeException("exportFormat", exportFormat, null);
            }
            return export.ExportResult(data, additionalData);
        }

        public class ModelRuntime
        {
            public string ClassName { get; set; }
            public string[] Properties { get; set; }
        }

        private static string GenerateClassFromJsonArray(string arr)
        {
            using (var ms = new MemoryStream())
            {
                using (var tw = new StreamWriter(ms))
                {
                    var hash = Math.Abs(arr.GetHashCode());

                    var mrj = new ModelRuntime();
                    mrj.ClassName = "Data" + hash;

                    var gen = new JsonClassGenerator();
                    gen.Example = arr;
                    gen.InternalVisibility = false;
                    gen.CodeWriter = new CSharpCodeWriter();
                    gen.ExplicitDeserialization = false;
                    gen.Namespace = null;

                    gen.NoHelperClass = true;
                    gen.SecondaryNamespace = null;
                    gen.TargetFolder = null;
                    gen.UseProperties = true;
                    gen.MainClass = mrj.ClassName;
                    gen.UsePascalCase = false;
                    gen.UseNestedClasses = true;
                    gen.ApplyObfuscationAttributes = false;
                    gen.SingleFile = true;
                    gen.ExamplesInDocumentation = false;
                    gen.OutputStream = tw;
                    gen.GenerateClasses();

                }
                return Encoding.ASCII.GetString(ms.ToArray());
            }
        }

        private static Type GenerateTypeFromProperties(string[] props)
        {
            var constructor = string.Join(",string ", props);
            var hash = Math.Abs(constructor.GetHashCode());

            var mrj = new ModelRuntime();
            mrj.ClassName = "Data" + hash;
            try
            {
                var typeExisting = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => a.GetTypes())
                    .FirstOrDefault(t => t.FullName.Equals(mrj.ClassName));

                if (typeExisting != null)
                    return typeExisting;

            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }
            mrj.Properties = props;

            var template = @"
using System;
public class @Model.ClassName {
//constructor
public @Model.ClassName (
    @foreach(var prop in Model.Properties){
    <text>string @prop , </text>
    }
    //add a fake property
string fake=null)
{
    @foreach(var prop in Model.Properties){
    <text>this.@prop = @prop;</text>
    }
}//end constructor
//properties
@foreach(var prop in Model.Properties){
    <text>public string @prop{get;set;}</text>
    }
 
}//end class               
";
            var code = Engine.Razor.RunCompile(template, mrj.ClassName, typeof(ModelRuntime), mrj);


            AssemblyCSharpBuilder builder = new("ExportCoreClass");
            builder.Domain = DomainManagement.Random;
            builder.Add(code);
            builder.OutputToFile = true;
            builder.OutputFolder = Path.Combine(Path.GetTempPath(), "ExportCoreClass");
            builder.GetAssembly();
            var ass = Path.Combine(builder.OutputFolder, "ExportCoreClass.dll");
            if (!File.Exists(ass))
            {
                throw new FileNotFoundException(ass);
            }
            var dd = Assembly.LoadFile(ass);
            var type = dd.DefinedTypes.First(t => t.Name == mrj.ClassName);
            return type;
        }

        private static Type GenerateTypeFromJson(string json)
        {

            var hash = Math.Abs(json.GetHashCode());

            var mrj = new ModelRuntime();
            mrj.ClassName = "Data" + hash;
            try
            {
                var typeExisting = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => a.GetTypes())
                    .FirstOrDefault(t => t.FullName.Equals(mrj.ClassName));

                if (typeExisting != null)
                    return typeExisting;

            }
            catch (Exception ex)
            {
                string s = ex.Message;
            }




            var code = GenerateClassFromJsonArray(json);

            AssemblyCSharpBuilder builder = new("ExportCoreClass");
            builder.Domain = DomainManagement.Random;
            builder.Add(code);
            builder.OutputToFile = true;
            builder.OutputFolder = Path.Combine(Path.GetTempPath(), "ExportCoreClass");
            builder.GetAssembly();

            var ass = Path.Combine(builder.OutputFolder, "ExportCoreClass.dll");
            if (!File.Exists(ass))
            {
                throw new FileNotFoundException(ass);
            }
            var dd = Assembly.LoadFile(ass);
            var type = dd.DefinedTypes.First(t => t.Name == mrj.ClassName);
            return type;
        }

        public static byte[] ExportDataJson(string jsonArray, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var list = FromJson(jsonArray);
            var type = list[0].GetType();
            return ExportDataWithType(list, exportFormat, type, additionalData);
        }

        static IList FromJson(string jsonArray)
        {
            var jObj = JArray.Parse(jsonArray);
            var type = GenerateTypeFromJson(jsonArray);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            // Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 0; i < jObj.Count; i++)
            {
                var item = jObj[i];

                dynamic obj = JsonConvert.DeserializeObject(item.ToString(), type);
                list.Add(obj);

            }
            return list as IList;

        }
        public static byte[] ExportDataDictionary(Dictionary<string, object>[] values, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var props = values[0].Keys.ToArray();
            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 0; i < values.Length; i++)
            {
                var item = values[i];
                var propsValue = item.Values.ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance,
                    null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);

        }
        public static byte[] ExportDataCsv(string[] csvWithHeader, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {

            var props = csvWithHeader[0].Split(new string[] { "," }, StringSplitOptions.None);
            if (props.Contains(""))
                throw new ArgumentException("header contains empty string");

            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            // Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 1; i < csvWithHeader.Length; i++)
            {
                var item = csvWithHeader[i];
                var propsValue = item.Split(new string[] { "," }, StringSplitOptions.None).ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance,
                    null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);
        }

        static internal IList IEnumerableFromDataTable(DataTable data)
        {
            var cols = data.Columns;

            var props = new string[cols.Count];
            for (int i = 0; i < cols.Count; i++)
            {
                props[i] = cols[i].ColumnName;
            }

            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            //in order to be found from Razor export
            Assembly.LoadFile(assembly.Location);
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType);
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var item = data.Rows[i];
                var propsValue = item.ItemArray.Select(it => it.ToString()).ToList();
                propsValue.Add("fake");
                dynamic obj = assembly.CreateInstance(type.FullName, true, BindingFlags.Public | BindingFlags.Instance,
                    null,
                    propsValue.ToArray(), null,
                    null);
                list.Add(obj);

            }
            return list as IList;

        }
        public static byte[] ExportDataFromDataTable(DataTable data, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var list = IEnumerableFromDataTable(data);
            var type = list[0].GetType();
            return ExportDataWithType(list as IEnumerable, exportFormat, type, additionalData);
        }

        public static byte[] ExportDataRSS(string rss, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {

            var json = GetDataFromRSS(rss);
            return ExportDataJson(json, exportFormat, additionalData);

        }

        static string GetDataFromRSS(string rss)
        {
            XDocument rssDoc;
            if (Uri.IsWellFormedUriString(rss, UriKind.Absolute))
            {
                rssDoc = XDocument.Load(rss);


            }
            else
            {
                rssDoc = XDocument.Parse(rss);
            }

            var data =
                rssDoc.Descendants("item")
                    .Select(
                        it =>
                            new
                            {
                                Title =
                                    (it.Element("title") ?? it.Element("description") ?? new XElement("noData")).Value,
                                Link = (it.Element("link") ?? new XElement("noData")).Value,
                                Description =
                                    (it.Element("description") ?? it.Element("title") ?? new XElement("noData")).Value
                            }).ToArray();
            var json = JsonConvert.SerializeObject(data);
            return json;
        }

        public static byte[] ExportDataIDataReader(IDataReader dr, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var dt = new DataTable();
            dt.Load(dr);
            return ExportDataFromDataTable(dt, exportFormat, additionalData);
        }

        public static byte[] ExportDataSet(DataSet ds, ExportToFormat exportFormat,
           params KeyValuePair<string, object>[] additionalData)
        {
            if (exportFormat != ExportToFormat.Excel2007)
                throw new ArgumentException("ready just for Excel 2007");

            var export = new ExportExcel2007<Tuple<string>>();
            return export.ExportMultipleSheets(ds);

        }

        public static byte[] ExportOpmlRSS(string dataOPML, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            if (exportFormat != ExportToFormat.Excel2007)
                throw new ArgumentException("ready just for Excel 2007");

            XDocument rssDoc;
            if (Uri.IsWellFormedUriString(dataOPML, UriKind.Absolute))
            {
                rssDoc = XDocument.Load(dataOPML);


            }
            else
            {
                rssDoc = XDocument.Parse(dataOPML);
            }

            var dataNodes = rssDoc.XPathSelectElements("//outline[@type='rss']").ToArray();
            var count = dataNodes.Length;
            var listData = new List<IList>(count);
            for (int i = 0; i < count; i++)
            {
                var rss = dataNodes[i].Attribute("xmlUrl").Value;
                var jsonArray = ExportFactory.GetDataFromRSS(rss);
                listData.Add(FromJson(jsonArray));
            }
            var export = new ExportExcel2007<Tuple<string>>();
            var data = export.ExportMultipleSheets(listData.ToArray(), additionalData);
            return data;

        }


    }
}