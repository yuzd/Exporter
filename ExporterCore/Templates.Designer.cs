﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码由工具生成。
//     运行时版本:4.0.30319.42000
//
//     对此文件的更改可能会导致不正确的行为，并且如果
//     重新生成代码，这些更改将会丢失。
// </auto-generated>
//------------------------------------------------------------------------------

namespace ExporterCore {
    using System;
    
    
    /// <summary>
    ///   一个强类型的资源类，用于查找本地化的字符串等。
    /// </summary>
    // 此类是由 StronglyTypedResourceBuilder
    // 类通过类似于 ResGen 或 Visual Studio 的工具自动生成的。
    // 若要添加或移除成员，请编辑 .ResX 文件，然后重新运行 ResGen
    // (以 /str 作为命令选项)，或重新生成 VS 项目。
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    public class Templates {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Templates() {
        }
        
        /// <summary>
        ///   返回此类使用的缓存的 ResourceManager 实例。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("ExporterCore.Templates", typeof(Templates).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   重写当前线程的 CurrentUICulture 属性，对
        ///   使用此强类型资源类的所有资源查找执行重写。
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&quot;1.0&quot;?&gt;
        ///&lt;?mso-application progid=&quot;Excel.Sheet&quot;?&gt;
        ///&lt;Workbook xmlns=&quot;urn:schemas-microsoft-com:office:spreadsheet&quot;
        ///xmlns:o=&quot;urn:schemas-microsoft-com:office:office&quot;
        ///xmlns:x=&quot;urn:schemas-microsoft-com:office:excel&quot;
        ///xmlns:ss=&quot;urn:schemas-microsoft-com:office:spreadsheet&quot;
        ///xmlns:html=&quot;http://www.w3.org/TR/REC-html40&quot;&gt;
        ///&lt;DocumentProperties xmlns=&quot;urn:schemas-microsoft-com:office:office&quot;&gt;
        /// &lt;Author&gt;Andrei Ignat&lt;/Author&gt;
        /// &lt;LastAuthor&gt;Andrei Ignat&lt;/LastAuthor&gt;
        /// &lt;Created&gt;$DateCreated;format=&quot;yyyy- [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        public static string Excel2003File {
            get {
                return ResourceManager.GetString("Excel2003File", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;Row&gt;
        ///@foreach(var item in Model){
        ///   &lt;Cell ss:StyleID=&apos;s21&apos;&gt;&lt;Data ss:Type=&apos;String&apos;&gt;@System.Security.SecurityElement.Escape(item)&lt;/Data&gt;&lt;/Cell&gt;
        ///} 
        ///&lt;/Row&gt; 的本地化字符串。
        /// </summary>
        public static string Excel2003Header {
            get {
                return ResourceManager.GetString("Excel2003Header", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;Row&gt;
        ///@foreach(var item in Model){	
        ///   &lt;Cell&gt;&lt;Data ss:Type=&apos;String&apos;&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/Data&gt;&lt;/Cell&gt;
        ///}&lt;/Row&gt; 的本地化字符串。
        /// </summary>
        public static string Excel2003Item {
            get {
                return ResourceManager.GetString("Excel2003Item", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;Worksheet ss:Name=&quot;Model.NameOfT&quot;&gt;
        /// &lt;Table  x:FullColumns=&quot;1&quot;
        ///  x:FullRows=&quot;1&quot;&gt;
        ///@Include(Model.NameOfT+&quot;Excel2003Header&quot;)
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;Excel2003Item&quot;,item)
        ///} 
        /// &lt;/Table&gt;
        /// &lt;WorksheetOptions xmlns=&quot;urn:schemas-microsoft-com:office:excel&quot;&gt;
        ///  &lt;Print&gt;
        ///   &lt;ValidPrinterInfo/&gt;
        ///   &lt;PaperSizeIndex&gt;9&lt;/PaperSizeIndex&gt;
        ///   &lt;HorizontalResolution&gt;600&lt;/HorizontalResolution&gt;
        ///   &lt;VerticalResolution&gt;600&lt;/VerticalResolution&gt;
        ///  &lt;/Print&gt;
        ///  &lt;Selected/&gt;
        ///  &lt;Panes&gt;
        ///   &lt;P [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        public static string Excel2003Sheet {
            get {
                return ResourceManager.GetString("Excel2003Sheet", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&apos;1.0&apos; encoding=&apos;UTF-8&apos; standalone=&apos;yes&apos; ?&gt;
        ///&lt;worksheet xmlns=&apos;http://schemas.openxmlformats.org/spreadsheetml/2006/main&apos; xmlns:r=&apos;http://schemas.openxmlformats.org/officeDocument/2006/relationships&apos;&gt;
        ///    &lt;sheetData&gt;
        ///
        ///@Include(Model.NameOfT+&quot;Excel2007Header&quot;)
        ///
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;Excel2007Item&quot;,item)
        ///} 
        ///
        ///    &lt;/sheetData&gt;
        ///&lt;/worksheet&gt; 的本地化字符串。
        /// </summary>
        public static string Excel2007File {
            get {
                return ResourceManager.GetString("Excel2007File", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;row&gt;
        ///@foreach(var item in Model){
        ///   &lt;c t=&apos;inlineStr&apos;&gt;
        ///                &lt;is&gt;
        ///                    &lt;t&gt;@System.Security.SecurityElement.Escape(item)&lt;/t&gt;
        ///                &lt;/is&gt;
        ///            &lt;/c&gt;
        ///} 
        ///&lt;/row&gt; 的本地化字符串。
        /// </summary>
        public static string Excel2007Header {
            get {
                return ResourceManager.GetString("Excel2007Header", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;row&gt;
        ///@foreach(var item in Model){
        ///   &lt;c t=&apos;inlineStr&apos;&gt;
        ///                &lt;is&gt;
        ///                    &lt;t&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())
        ///                    &lt;/t&gt;
        ///                &lt;/is&gt;
        ///    &lt;/c&gt;
        ///   }
        ///&lt;/row&gt; 的本地化字符串。
        /// </summary>
        public static string Excel2007Item {
            get {
                return ResourceManager.GetString("Excel2007Item", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;html&gt;
        ///    &lt;head&gt;
        ///    &lt;title&gt;Export&lt;/title&gt;
        ///    &lt;/head&gt;
        ///    &lt;body&gt;
        ///    
        ///&lt;Table border=&quot;1&quot;&gt;
        ///
        ///@Include(Model.NameOfT+&quot;HtmlHeader&quot;)
        ///
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;HtmlItem&quot;,item)
        ///} 
        ///
        /// &lt;/Table&gt;
        ///                
        ///         Generated on @ViewBag.DateCreated
        ///    &lt;/body&gt;
        ///&lt;/html&gt; 的本地化字符串。
        /// </summary>
        public static string HtmlFile {
            get {
                return ResourceManager.GetString("HtmlFile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;tr&gt;
        ///@foreach(var item in Model){
        ///   &lt;th&gt;@System.Security.SecurityElement.Escape(item)&lt;/th&gt;
        ///} 
        ///&lt;/tr&gt; 的本地化字符串。
        /// </summary>
        public static string HtmlHeader {
            get {
                return ResourceManager.GetString("HtmlHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;tr&gt;
        ///@foreach(var item in Model){
        ///   &lt;td&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/td&gt;
        ///} 
        ///&lt;/tr&gt; 的本地化字符串。
        /// </summary>
        public static string HtmlItem {
            get {
                return ResourceManager.GetString("HtmlItem", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;itext author=&quot;Andrei Ignat&quot; title=&quot;Collection&quot;&gt;
        ///
        ///&lt;chapter numberdepth=&quot;0&quot;&gt;
        ///&lt;newline /&gt;
        ///&lt;section numberdepth=&quot;0&quot;&gt;
        ///&lt;table width=&quot;100%&quot;  cellspacing=&quot;0&quot; cellpadding=&quot;2&quot; columns=&quot;$ItemToDisplay.Properties.keys.Count$&quot; grayfill=&quot;0.90&quot;&gt;
        ///
        ///@Include(Model.NameOfT+&quot;iTextSharp4Header&quot;)
        ///
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;iTextSharp4Item&quot;,item)
        ///}     
        ///&lt;/table&gt;
        ///&lt;/section&gt;
        ///
        ///&lt;/chapter&gt;
        ///
        ///
        ///
        ///&lt;/itext&gt; 的本地化字符串。
        /// </summary>
        public static string iTextSharp4File {
            get {
                return ResourceManager.GetString("iTextSharp4File", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;row&gt;
        ///@foreach(var item in Model){
        ///&lt;cell&gt;&lt;phrase font=&apos;Arial&apos; size=&apos;12&apos; style=&apos;bold&apos;&gt;@System.Security.SecurityElement.Escape(item)&lt;/phrase&gt;&lt;/cell&gt;
        ///}
        ///&lt;/row&gt; 的本地化字符串。
        /// </summary>
        public static string iTextSharp4Header {
            get {
                return ResourceManager.GetString("iTextSharp4Header", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;row&gt;
        ///@foreach(var item in Model){
        ///    &lt;cell&gt;&lt;phrase font=&apos;Times New Roman&apos; size=&apos;8&apos;&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/phrase&gt;&lt;/cell&gt;
        ///}
        ///&lt;/row&gt; 的本地化字符串。
        /// </summary>
        public static string iTextSharp4Item {
            get {
                return ResourceManager.GetString("iTextSharp4Item", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&apos;1.0&apos; encoding=&apos;UTF-8&apos;?&gt;
        ///&lt;office:document-content xmlns:office=&apos;urn:oasis:names:tc:opendocument:xmlns:office:1.0&apos; xmlns:style=&apos;urn:oasis:names:tc:opendocument:xmlns:style:1.0&apos; xmlns:text=&apos;urn:oasis:names:tc:opendocument:xmlns:text:1.0&apos; xmlns:table=&apos;urn:oasis:names:tc:opendocument:xmlns:table:1.0&apos; xmlns:draw=&apos;urn:oasis:names:tc:opendocument:xmlns:drawing:1.0&apos; xmlns:fo=&apos;urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0&apos; xmlns:xlink=&apos;http://www.w3.org/1999/xlink&apos; xmlns:dc=&apos;http://purl. [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        public static string ODSFile {
            get {
                return ResourceManager.GetString("ODSFile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;table:table-column table:style-name=&apos;Standard&apos;  table:number-columns-repeated=&apos;@Model.Length&apos;/&gt; 
        ///&lt;table:table-row&gt; 
        ///@foreach(var item in Model){
        ///	&lt;table:table-cell table:style-name=&apos;Standard&apos; office:value-type=&apos;string&apos;&gt;
        ///		&lt;text:p&gt;@System.Security.SecurityElement.Escape(item)&lt;/text:p&gt;
        ///	&lt;/table:table-cell&gt;
        ///}
        ///&lt;/table:table-row&gt; 的本地化字符串。
        /// </summary>
        public static string ODSHeader {
            get {
                return ResourceManager.GetString("ODSHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;table:table-row&gt;
        ///@foreach(var item in Model){
        ///   
        ///   	&lt;table:table-cell table:style-name=&apos;Standard&apos;  office:value-type=&apos;string&apos;&gt;
        ///		  &lt;text:p&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/text:p&gt;
        ///&lt;/table:table-cell&gt;                    
        ///   }
        ///&lt;/table:table-row&gt; 的本地化字符串。
        /// </summary>
        public static string ODSItem {
            get {
                return ResourceManager.GetString("ODSItem", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&apos;1.0&apos; encoding=&apos;UTF-8&apos;?&gt;
        ///&lt;office:document-content xmlns:office=&apos;urn:oasis:names:tc:opendocument:xmlns:office:1.0&apos; xmlns:style=&apos;urn:oasis:names:tc:opendocument:xmlns:style:1.0&apos; xmlns:text=&apos;urn:oasis:names:tc:opendocument:xmlns:text:1.0&apos; xmlns:table=&apos;urn:oasis:names:tc:opendocument:xmlns:table:1.0&apos; xmlns:draw=&apos;urn:oasis:names:tc:opendocument:xmlns:drawing:1.0&apos; xmlns:fo=&apos;urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0&apos; xmlns:xlink=&apos;http://www.w3.org/1999/xlink&apos; xmlns:dc=&apos;http://purl.o [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        public static string ODTFile {
            get {
                return ResourceManager.GetString("ODTFile", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;table:table-column table:style-name=&apos;Standard&apos;  table:number-columns-repeated=&apos;@Model.Length&apos;/&gt; 
        ///&lt;table:table-row&gt; 
        ///@foreach(var item in Model){
        ///               &lt;table:table-cell table:style-name=&apos;Standard&apos;
        ///		  office:value-type=&apos;string&apos;&gt;
        ///
        ///		  	&lt;text:p&gt;
        ///                                      @System.Security.SecurityElement.Escape(item)
        ///                              &lt;/text:p&gt;
        ///		
        ///                &lt;/table:table-cell&gt;
        ///}
        ///&lt;/table:table-row&gt; 的本地化字符串。
        /// </summary>
        public static string ODTHeader {
            get {
                return ResourceManager.GetString("ODTHeader", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;table:table-row&gt; 
        ///@foreach(var item in Model){
        ///&lt;table:table-cell table:style-name=&apos;Standard&apos;
        /// office:value-type=&apos;string&apos;&gt;
        ///		  	
        ///&lt;text:p&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/text:p&gt;
        ///
        ///&lt;/table:table-cell&gt;   
        ///}
        ///&lt;/table:table-row&gt; 的本地化字符串。
        /// </summary>
        public static string ODTItem {
            get {
                return ResourceManager.GetString("ODTItem", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&quot;1.0&quot;?&gt;
        ///&lt;w:wordDocument xmlns:w=&quot;http://schemas.microsoft.com/office/word/2003/wordml&quot; &gt;        
        ///    &lt;w:body&gt;
        ///        &lt;w:tbl&gt;
        ///    
        ///@Include(Model.NameOfT+&quot;Word2003Header&quot;)
        ///
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;Word2003Item&quot;,item)
        ///} 
        ///
        ///        &lt;/w:tbl&gt;
        ///    &lt;/w:body&gt;
        ///    
        ///&lt;/w:wordDocument&gt; 的本地化字符串。
        /// </summary>
        public static string Word2003File {
            get {
                return ResourceManager.GetString("Word2003File", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;w:tr&gt;
        ///@foreach(var item in Model){
        ///           &lt;w:tc&gt;
        ///                &lt;w:p&gt;
        ///                    &lt;w:r&gt;
        ///                        &lt;w:rPr&gt;
        ///                            &lt;w:b w:val=&apos;on&apos;/&gt;
        ///                            &lt;w:t&gt;
        ///                                @System.Security.SecurityElement.Escape(item)
        ///                            &lt;/w:t&gt;
        ///                        &lt;/w:rPr&gt;
        ///                    &lt;/w:r&gt;
        ///                &lt;/w:p&gt;
        ///            &lt;/w:tc&gt;                
        ///}
        ///&lt;/w:tr&gt; 的本地化字符串。
        /// </summary>
        public static string Word2003Header {
            get {
                return ResourceManager.GetString("Word2003Header", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;w:tr&gt;
        ///@foreach(var item in Model){
        ///   &lt;w:tc&gt;
        ///    &lt;w:p&gt;
        ///        &lt;w:r&gt;
        ///            &lt;w:t&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/w:t&gt;
        ///        &lt;/w:r&gt;
        ///    &lt;/w:p&gt;
        ///    &lt;/w:tc&gt;
        ///   }
        ///&lt;/w:tr&gt; 的本地化字符串。
        /// </summary>
        public static string Word2003Item {
            get {
                return ResourceManager.GetString("Word2003Item", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;?xml version=&quot;1.0&quot; encoding=&quot;UTF-8&quot; standalone=&quot;yes&quot;?&gt;
        ///&lt;w:document xmlns:w=&quot;http://schemas.openxmlformats.org/wordprocessingml/2006/main&quot;&gt;
        ///    &lt;w:body&gt;
        ///        &lt;w:tbl&gt;
        ///
        ///@Include(Model.NameOfT+&quot;Word2007Header&quot;)
        ///
        ///@foreach(var item in Model.Data){
        ///  @Include(Model.NameOfT+&quot;Word2007Item&quot;,item)
        ///} 
        ///        &lt;/w:tbl&gt;
        ///    &lt;/w:body&gt;
        ///&lt;/w:document&gt; 的本地化字符串。
        /// </summary>
        public static string Word2007File {
            get {
                return ResourceManager.GetString("Word2007File", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;w:tr&gt;
        ///@foreach(var item in Model){
        ///                &lt;w:tc&gt;
        ///                    &lt;w:p&gt;
        ///                        &lt;w:r&gt;
        ///                            &lt;w:rPr&gt;
        ///                                &lt;w:b w:val=&apos;on&apos;/&gt;
        ///                                &lt;w:t&gt;
        ///                                      @System.Security.SecurityElement.Escape(item)
        ///                                &lt;/w:t&gt;
        ///                            &lt;/w:rPr&gt;
        ///                        &lt;/w:r&gt;
        ///                    &lt;/w:p&gt;
        ///                &lt;/w:tc&gt;
        ///}
        ///            &lt;/ [字符串的其余部分被截断]&quot;; 的本地化字符串。
        /// </summary>
        public static string Word2007Header {
            get {
                return ResourceManager.GetString("Word2007Header", resourceCulture);
            }
        }
        
        /// <summary>
        ///   查找类似 &lt;w:tr&gt;
        ///@foreach(var item in Model){
        ///   &lt;w:tc&gt;
        ///    &lt;w:p&gt;
        ///        &lt;w:r&gt;
        ///              &lt;w:t&gt;@@System.Security.SecurityElement.Escape((((object)Model.@item) ?? &quot;&quot;).ToString())&lt;/w:t&gt;
        ///        &lt;/w:r&gt;
        ///    &lt;/w:p&gt;
        ///    &lt;/w:tc&gt;
        ///   }
        ///&lt;/w:tr&gt; 的本地化字符串。
        /// </summary>
        public static string Word2007Item {
            get {
                return ResourceManager.GetString("Word2007Item", resourceCulture);
            }
        }
    }
}