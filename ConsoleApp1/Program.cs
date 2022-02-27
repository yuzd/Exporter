using System.Diagnostics;
using ExportImplementation;

//var arrCSV = new List<string>();
//arrCSV.Add("Name,WebSite,连接");
//arrCSV.Add("111,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
//arrCSV.Add("123,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

//var data = ExportFactory.ExportDataCsv(arrCSV.ToArray(), ExportToFormat.Excel2007);
//File.WriteAllBytes("a.xlsx", data);


//string json = @"[
//        { 'Name':'Andrei Ignat', 
//            'WebSite':'http://msprogrammer.serviciipeweb.ro/',
//            'CV':'http://serviciipeweb.ro/iafblog/content/binary/cv.doc'        
//        },
//    { 'Name':'Your Name', 
//            'WebSite':'http://your website',
//            'CV':'cv.doc'        
//        }
//    ]";
//var data2 = ExportFactory.ExportDataJson(json, ExportToFormat.Excel2007);
//File.WriteAllBytes("a.xlsx", data2);


List<Person> listWithPerson = new List<Person>
{
    new Person
    {
        Name = "aa",
        Aget = 12
    },
    new Person
    {
        Name = "dasda",
        Aget = 1222
    }
};
var data = ExportFactory.ExportData(listWithPerson, ExportToFormat.Excel2007);
File.WriteAllBytes("a.xlsx", data);


public class Person
{
    public string Name { get; set; }
    public int Aget { get; set; }
}