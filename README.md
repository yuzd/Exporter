# nuget

Install-Package ExporterCore


# csv(逗号分隔)导出为excel
```csharp

var arrCSV = new List<string>();
arrCSV.Add("Name,WebSite,连接");
arrCSV.Add("111,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");
arrCSV.Add("123,http://msprogrammer.serviciipeweb.ro/,http://serviciipeweb.ro/iafblog/content/binary/cv.doc");

var data = ExportFactory.ExportDataCsv(arrCSV.ToArray(), ExportToFormat.Excel2007);
File.WriteAllBytes("a.xlsx", data);

```

# json导出为excel
```csharp
string json = @"[
        { 'Name':'Andrei Ignat', 
            'WebSite':'http://msprogrammer.serviciipeweb.ro/',
            'CV':'http://serviciipeweb.ro/iafblog/content/binary/cv.doc'        
        },
    { 'Name':'Your Name', 
            'WebSite':'http://your website',
            'CV':'cv.doc'        
        }
    ]";
var data2 = ExportFactory.ExportDataJson(json, ExportToFormat.Excel2007);
File.WriteAllBytes("a.xlsx", data2);

```

# list导出为excel
```csharp

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

```

# 多个list导出同个excel的多张表

```csharp

var p = new Person { Name = "andrei", WebSite = "http://msprogrammer.serviciipeweb.ro/", CV = "http://serviciipeweb.ro/iafblog/content/binary/cv.doc" };
var p1 = new Person { Name = "you", WebSite = "http://yourwebsite.com/" };
var list = new List<Person>() { p, p1 };

var kvp = new List<Tuple<string, string>>();
for (int i = 0; i < 10; i++)
{
    var q = new Tuple<string, string>("This is key " + i, "Value " + i);
    kvp.Add(q);
}

var export = new ExportExcel2007<Person>();
var data = export.ExportMultipleSheets(new IList[] { list, kvp });
File.WriteAllBytes("multiple.xlsx", data);
```
