using System.Collections;
using System.Data;
using System.Reflection;
using System.Reflection.Emit;
using LamarCompiler;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RazorEngineCore;
using Xamasoft.JsonClassGenerator;
using Xamasoft.JsonClassGenerator.CodeWriters;
using Encoding = System.Text.Encoding;

namespace ExporterCore
{
    public static class ExportFactory
    {

        /// <summary>
        /// 根据属性生成代码模板
        /// </summary>
        private static readonly IRazorEngineCompiledTemplate PropertyClassTemplateCompiled;


        static ExportFactory()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            PropertyClassTemplateCompiled = razorEngine.Compile(Templates.ClassTemplate);
        }

        /// <summary>
        /// 根据不同的类型创建不同的导出对象
        /// </summary>
        /// <param name="data"></param>
        /// <param name="exportFormat"></param>
        /// <param name="type"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static byte[]? ExportDataWithType(IEnumerable data, ExportToFormat exportFormat, Type type,
            params KeyValuePair<string, object>[] additionalData)

        {
            Type exportType;
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
                case ExportToFormat.Word:
                    exportType = typeof(ExportWord<>).MakeGenericType(type);
                    break;
                case ExportToFormat.Excel:
                    exportType = typeof(ExportExcel<>).MakeGenericType(type);
                    break;
                case ExportToFormat.JSON:
                    exportType = typeof(ExportJson<>).MakeGenericType(type);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
            }

            dynamic? export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export?.ExportResult(data1, additionalData);
        }

        /// <summary>
        /// 根据泛型导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
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
                case ExportToFormat.Word:
                    export = new ExportWord<T>();
                    break;
                case ExportToFormat.Excel:
                    export = new ExportExcel<T>();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(exportFormat), exportFormat, null);
            }
            return export.ExportResult(data, additionalData);
        }

        /// <summary>
        /// 根据jsonarray导出
        /// </summary>
        /// <param name="jsonArray"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public static byte[]? ExportDataJson(string jsonArray, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var list = FromJson(jsonArray);
            if (list == null)
            {
                return null;
            }
            var type = list?[0]?.GetType();
            if (type == null)
            {
                return null;
            }
            return ExportDataWithType(list!, exportFormat, type, additionalData);
        }


        /// <summary>
        /// 将字典集合导出
        /// </summary>
        /// <param name="values"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public static byte[]? ExportDataDictionary(Dictionary<string, object>[] values, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var props = values[0].Keys.ToArray();
            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType)!;
            foreach (var item in values)
            {
                var propsValue = item.Values.ToList();
                propsValue.Add("fake");
                dynamic? obj = assembly.CreateInstance(type.FullName!,
                    true,
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    propsValue.ToArray(),
                    null,
                    null);
                list.Add(obj);
            }
            return ExportDataWithType((list as IEnumerable)!, exportFormat, type, additionalData);

        }

        /// <summary>
        /// 将csv导出
        /// </summary>
        /// <param name="csvWithHeader"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static byte[]? ExportDataCsv(string[] csvWithHeader, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {

            var props = csvWithHeader[0].Split(new string[] { "," }, StringSplitOptions.None);
            if (props.Contains(""))
                throw new ArgumentException("header contains empty string");

            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType)!;
            for (int i = 1; i < csvWithHeader.Length; i++)
            {
                var item = csvWithHeader[i];
                var propsValue = item.Split(new string[] { "," }, StringSplitOptions.None).ToList();
                propsValue.Add("fake");
                dynamic? obj = assembly.CreateInstance(type.FullName!,
                    true,
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    // ReSharper disable once CoVariantArrayConversion
                    propsValue.ToArray(),
                    null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType((list as IEnumerable)!, exportFormat, type, additionalData);
        }
        public static byte[]? ExportDataCsv(IEnumerable<string[]> csvWithHeader, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var stringsEnumerable = csvWithHeader as string[][] ?? csvWithHeader.ToArray();
            var props = stringsEnumerable[0];
            if (props.Contains(""))
                throw new ArgumentException("header contains empty string");

            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType)!;
            for (int i = 1; i < stringsEnumerable.Length; i++)
            {
                var item = stringsEnumerable[i];
                if(props.Length!=item.Length)continue;
                
                var propsValue = item.ToList();
                propsValue.Add("fake");
                dynamic? obj = assembly.CreateInstance(type.FullName!,
                    true,
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    // ReSharper disable once CoVariantArrayConversion
                    propsValue.ToArray(),
                    null,
                    null);
                list.Add(obj);

            }
            return ExportDataWithType((list as IEnumerable)!, exportFormat, type, additionalData);
        }

        /// <summary>
        /// 将DataTable导出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public static byte[]? ExportDataFromDataTable(DataTable data, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var list = EnumerableFromDataTable(data);
            var type = list?[0]?.GetType();
            if (type == null)
            {
                return null;
            }
            return ExportDataWithType(list!, exportFormat, type, additionalData);
        }


        /// <summary>
        /// 将DataReader导出
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public static byte[]? ExportDataIDataReader(IDataReader dr, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            var dt = new DataTable();
            dt.Load(dr);
            return ExportDataFromDataTable(dt, exportFormat, additionalData);
        }


        /// <summary>
        /// 将DataSet导出Excel xlsx格式
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="exportFormat"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static byte[]? ExportDataSet(DataSet ds, ExportToFormat exportFormat,
            params KeyValuePair<string, object>[] additionalData)
        {
            if (exportFormat != ExportToFormat.Excel)
                throw new ArgumentException("ready just for Excel 2007");

            var export = new ExportExcel<Tuple<string>>();
            return export.ExportMultipleSheets(ds);

        }

        /// <summary>
        /// 将DataTable转化成List对象
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        internal static IList? EnumerableFromDataTable(DataTable data)
        {
            var cols = data.Columns;

            var props = new string[cols.Count];
            for (int i = 0; i < cols.Count; i++)
            {
                props[i] = cols[i].ColumnName;
            }

            var type = GenerateTypeFromProperties(props);
            var assembly = type.Assembly;
            var listType = typeof(List<>).MakeGenericType(type);

            dynamic list = Activator.CreateInstance(listType)!;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                var item = data.Rows[i];
                var propsValue = item.ItemArray.Select(it => it!.ToString()).ToList();
                propsValue.Add("fake");
                dynamic? obj = assembly.CreateInstance(type.FullName!,
                    true,
                    BindingFlags.Public | BindingFlags.Instance,
                    null,
                    // ReSharper disable once CoVariantArrayConversion
                    propsValue.ToArray()!,
                    null,
                    null);
                list.Add(obj);

            }
            return list as IList;

        }


        /// <summary>
        /// 从jsonarray里面创建对象的申明模板
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        private static string GenerateClassFromJsonArray(string arr)
        {
            using var ms = new MemoryStream();
            using var tw = new StreamWriter(ms);
            var hash = Math.Abs(arr.GetHashCode());

            var mrj = new ModelRuntime
            {
                ClassName = "Data" + hash
            };

            var gen = new JsonClassGenerator
            {
                Example = arr,
                InternalVisibility = false,
                CodeWriter = new CSharpCodeWriter(),
                ExplicitDeserialization = false,
                Namespace = null,
                NoHelperClass = true,
                SecondaryNamespace = null,
                TargetFolder = null,
                UseProperties = true,
                MainClass = mrj.ClassName,
                UsePascalCase = false,
                UseNestedClasses = true,
                ApplyObfuscationAttributes = false,
                SingleFile = true,
                ExamplesInDocumentation = false,
                OutputStream = tw
            };

            gen.GenerateClasses();
            gen.OutputStream.Flush();
            return Encoding.ASCII.GetString(ms.ToArray());
        }

        /// <summary>
        /// 检测当前domain已经创建好了相同的class
        /// </summary>
        /// <param name="className"></param>
        /// <returns></returns>
        private static Type? GetExistedTypeInCurrentDomain(string className)
        {
            try
            {
                // 检测当前domain已经创建好了相同的class
                var typeExisting = AppDomain.CurrentDomain.GetAssemblies()
                    .SelectMany(a => a.GetTypes())
                    .FirstOrDefault(t => t.FullName != null && t.FullName.Equals(className));

                if (typeExisting != null)
                    return typeExisting;

            }
            catch (Exception)
            {
                //ignore
            }

            return null;
        }

        /// <summary>
        /// 根据属性创建class
        /// </summary>
        /// <param name="props"></param>
        /// <returns></returns>
        private static Type GenerateTypeFromProperties(string[] props)
        {
            var constructor = string.Join(",string ", props);
            var hash = Math.Abs(constructor.GetHashCode());
            var className = "Data" + hash;
            var existedType = GetExistedTypeInCurrentDomain(className);
            if (existedType != null)
            {
                return existedType;
            }

            var mrj = new ModelRuntime
            {
                ClassName = className,
                Properties = props
            };

            var code = PropertyClassTemplateCompiled.Run(mrj);
            var generator = new AssemblyGenerator();
            var assembly = generator.Generate(code);
            var type = assembly.GetExportedTypes().Single();
            return type;
        }

        /// <summary>
        /// 根据json创建class
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        private static Type GenerateTypeFromJson(string json)
        {

            var hash = Math.Abs(json.GetHashCode());
            var className = "Data" + hash;
            var existedType = GetExistedTypeInCurrentDomain(className);
            if (existedType != null)
            {
                return existedType;
            }


            var code = GenerateClassFromJsonArray(json);
            var generator = new AssemblyGenerator();
            var assembly = generator.Generate(code);
            var type = assembly.GetExportedTypes().Single();
            return type;
        }


        /// <summary>
        /// jsonArray转成List对象
        /// </summary>
        /// <param name="jsonArray"></param>
        /// <returns></returns>
        private static IList? FromJson(string jsonArray)
        {
            var jObj = JArray.Parse(jsonArray);
            if (jObj == null || !jObj.Any())
            {
                return null;
            }
            var type = GenerateTypeFromJson(jsonArray);
            var listType = typeof(List<>).MakeGenericType(type);
            dynamic? list = Activator.CreateInstance(listType);
            foreach (var item in jObj)
            {
                dynamic obj = JsonConvert.DeserializeObject(item.ToString(), type);
                list?.Add(obj);
            }
            return list as IList;

        }




        public class ModelRuntime
        {
            /// <summary>
            /// class名称
            /// </summary>
            public string ClassName { get; set; } = "";
            /// <summary>
            /// 属性列表
            /// </summary>
            public string[]? Properties { get; set; }
        }

    }
}