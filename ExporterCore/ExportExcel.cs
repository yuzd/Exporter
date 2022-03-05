using System.Collections;
using System.Data;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using RazorEngineCore;

namespace ExporterCore
{
    public class ExportExcel<T> : Export<T> where T : class
    {

        private const string Excel2007Header = "Excel2007Header";
        private const string Excel2007Item = "Excel2007Item";



        public ExportExcel()
        {
            var props = properties.Select(it => it.Name).ToArray();
            ExportCollection = Templates.Excel2007File;
            ExportHeader = Excel2007RazorTemplate.ExportHeaderCompiled.Run(props);
            ExportItem = Excel2007RazorTemplate.ExportItemCompiled.Run(props);
        }


        /// <summary>
        /// 导出excel xlsx格式
        /// </summary>
        /// <param name="data"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public override byte[] ExportResult(List<T> data, params KeyValuePair<string, object>[] additionalData)
        {
            string result = ExportResultStringPart(data);
            return CreateExcel2007(new[] { TType.Name }, new[] { result });
        }


        /// <summary>
        /// 生成单个sheet的excel内容模板
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public string ExportResultStringPart(List<T> data)
        {
            var modelTemplate = new ModelTemplate<T>(data);
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();

            IDictionary<string, string> parts = new Dictionary<string, string>()
            {
                {TType.Name + Excel2007Header, ExportHeader},
                {TType.Name + Excel2007Item, ExportItem}
            };

            IncludeCompiledTemplate compiledTemplate = razorEngine.Compile(ExportCollection, parts);
            string result = compiledTemplate.Run(modelTemplate);
            return result;
        }

        /// <summary>
        /// 将List数组导出多个sheet
        /// </summary>
        /// <param name="data"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        public byte[]? ExportMultipleSheets(IList[]? data, params KeyValuePair<string, object>[] additionalData)
        {
            if (data == null)
                return null;
            if (data.Length == 0)
                return null;

            var result = new string[data.Length];
            var names = new string[data.Length];
            for (int i = 0; i < data.Length; i++)
            {
                var list = data[i];
                if (list.Count == 0)
                    continue;
                names[i] = list?[0]!.GetType()?.Name ?? "";
                result[i] = ExportResultStringPartNotGeneric(list!, additionalData);

            }
            return CreateExcel2007(names, result);
        }



        /// <summary>
        /// 将单个List导出excel的单个sheet
        /// </summary>
        /// <param name="data"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public string ExportResultStringPartNotGeneric(IList data, params KeyValuePair<string, object>[] additionalData)
        {
            if (data.Count < 1)
            {
                throw new ArgumentException(nameof(data));
            }
            var type = data[0]!.GetType();
            dynamic exportType = typeof(ExportExcel<>).MakeGenericType(type);
            dynamic export = Activator.CreateInstance(exportType);
            dynamic data1 = data;
            return export.ExportResultStringPart(data1, additionalData);
        }


        /// <summary>
        /// DataSet导出多个excelsheet
        /// </summary>
        /// <param name="data"></param>
        /// <param name="additionalData"></param>
        /// <returns></returns>
        internal byte[]? ExportMultipleSheets(DataSet data, params KeyValuePair<string, object>[] additionalData)
        {
            if (data.Tables.Count == 0)
                return null;

            var result = new string[data.Tables.Count];
            var names = new string[data.Tables.Count];
            for (int i = 0; i < data.Tables.Count; i++)
            {
                var table = data.Tables[i];
                if (table.Rows.Count == 0)
                    continue;
                names[i] = table.TableName;
                if (string.IsNullOrWhiteSpace(names[i]))
                {
                    names[i] = "DataTable" + i;
                }
                var list = ExportFactory.EnumerableFromDataTable(table);
                if (list == null)
                {
                    return null;
                }
                result[i] = ExportResultStringPartNotGeneric(list, additionalData);

            }
            return CreateExcel2007(names, result);
        }

        /// <summary>
        /// 生成excel字节数组
        /// </summary>
        /// <param name="worksheetName"></param>
        /// <param name="textSheet"></param>
        /// <returns></returns>
        private byte[] CreateExcel2007(string[] worksheetName, string[] textSheet)
        {
            using var ms = new MemoryStream();
            using var sd = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            var workbook = sd.AddWorkbookPart();
            var strSheets = "<sheets>";
            for (var i = 0; i < worksheetName.Length; i++)
            {


                var sheet = workbook.AddNewPart<WorksheetPart>();

                WriteToPart(sheet, textSheet[i]);

                strSheets += string.Format("<sheet name=\"{1}\" sheetId=\"{2}\" r:id=\"{0}\" />",
                    workbook.GetIdOfPart(sheet), worksheetName[i], (i + 1));


            }
            strSheets += "</sheets>";
            WriteToPart(workbook, string.Format(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">{0}</workbook>",
                strSheets
            ));

            sd.Close();
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


    public class Excel2007RazorTemplate
    {
        /// <summary>
        /// excel的每一项
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportItemCompiled;

        /// <summary>
        /// excel的头部
        /// </summary>
        internal static readonly IRazorEngineCompiledTemplate ExportHeaderCompiled;

        static Excel2007RazorTemplate()
        {
            IRazorEngine razorEngine = new RazorEngineCore.RazorEngine();
            ExportHeaderCompiled = razorEngine.Compile(Templates.Excel2007Header);
            ExportItemCompiled = razorEngine.Compile(Templates.Excel2007Item);
        }
    }
}