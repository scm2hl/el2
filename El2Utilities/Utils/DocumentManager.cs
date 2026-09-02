using El2Core.Models;
using Microsoft.Extensions.Logging;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace El2Core.Utils
{
    public interface IDocument
    {
    }
    /// <summary>
    /// Represents a document with various parts and their corresponding values.
    /// </summary>
    public abstract class Document()
    {
        private readonly Dictionary<DocumentPart, string> parts =[];
        public string this[DocumentPart key]
        {
            get => parts[key];
            set => parts[key] = value;
        }
        public int Count => parts.Count;
        public HashSet<DocumentPart> Keys => [.. parts.Keys];
    }
    public class FirstPartDocument : Document { }
    public class VmpbDocument : Document { }
    public class WorkAreaDocument : Document { }
    /// <summary>
    /// Represents the information of a document, including its type, root path, template, and other relevant parts.
    /// </summary>
    /// <param name="container"></param>
    public abstract class DocumentInfo(IContainerExtension container)
    {
        private IContainerExtension Container => container;
        public abstract Document CreateDocumentInfos();
        public abstract Document CreateDocumentInfos(string[] folders, bool isDummy);
        public abstract Document GetDocument();
        public void SaveDocumentData(Document document)
        {
            using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            Rule rule;
            if(db.Rules.All(x => x.RuleValue.Trim() != document[DocumentPart.Type]))
            {
                rule = new Rule()
                { RuleValue = document[DocumentPart.Type], RuleName = document[DocumentPart.Type] };
                db.Rules.Add(rule);
            }
            else rule = db.Rules.First(x => x.RuleValue == document[DocumentPart.Type]);

            StringWriter sw = new();
            Serialize(sw, document);
            rule.RuleData = sw.ToString();

            db.SaveChanges();
        }
        /// <summary>
        /// Collects the document data and creates the necessary directories based on the SavePath in the document.
        /// </summary>
        /// <param name="document"></param>
        public void Collect(Document document)
        {
            if (string.IsNullOrEmpty(document[DocumentPart.SavePath])) { return; }
            if (!Directory.Exists(document[DocumentPart.RootPath])) return;
            string path = document[DocumentPart.RootPath];
            foreach(var s in document[DocumentPart.SavePath].Split(Path.DirectorySeparatorChar))
            {
                path = Path.Combine(path, s);
                if(!Directory.Exists(path)) Directory.CreateDirectory(path);
            }
        }
        private static void Serialize(TextWriter writer, Document dictionary)
        {

            List<Entry> entries = [];

            foreach(var k in dictionary.Keys)
            {
                entries.Add(new Entry(k, dictionary[k]));
            }

            var serializer = XmlSerializerHelper.GetSerializer(typeof(List<Entry>));
            serializer.Serialize(writer, entries);
        }
    }
    /// <summary>
    /// Represents the information of a measure first part document.
    /// </summary>
    public class MeasureFirstPartInfo : DocumentInfo
    {
        private Document document;
        ILogger logger;

        public MeasureFirstPartInfo(IContainerExtension container) : base(container)
        {
            var factory = container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<MeasureFirstPartInfo>();
        }

        public override Document CreateDocumentInfos(string[]? folders, bool isDummy = false)
        {
            try
            {
                document = new FirstPartDocument();
                document[DocumentPart.Type] = "FirstPart";
                document[DocumentPart.RootPath] = string.Empty;
                document[DocumentPart.Template] = string.Empty;
                document[DocumentPart.MaterialRegularEx] = string.Empty;
                document[DocumentPart.JumpTarget] = string.Empty;
                document[DocumentPart.RasterFolder1] = string.Empty;
                if (RuleInfo.Rules.Keys.Contains(document[DocumentPart.Type]) == false) return document;
                var xml = XmlSerializerHelper.GetSerializer(typeof(List<Entry>));

                TextReader reader = new StringReader(RuleInfo.Rules[document[DocumentPart.Type]].RuleData);
                List<Entry> doc = (List<Entry>)xml.Deserialize(reader);
                foreach (var entry in doc)
                {
                    DocumentPart DokuPart = (DocumentPart)Enum.Parse(typeof(DocumentPart), entry.Key.ToString());
                    document[DokuPart] = (string)entry.Value;
                }
                if (folders != null)
                {

                    document[DocumentPart.TTNR] = folders[0];
                    Regex regex = new Regex(document[DocumentPart.MaterialRegularEx]);
                    Match match2 = regex.Match(folders[0]);
                    StringBuilder nsb = new();
                    foreach (Group ma in match2.Groups.Values.Skip(1))
                    {
                        if (ma.Value != folders[0])
                        {
                            nsb.Append(ma.Value).Append(Path.DirectorySeparatorChar);
                        }
                    }
                    document[DocumentPart.SavePath] = nsb.ToString();
                    FileInfo f = new(document[DocumentPart.Template]);
                    document[DocumentPart.File] = Path.Combine(
                        document[DocumentPart.RootPath],
                        document[DocumentPart.SavePath],
                        folders[0] + "_1.Gutteil" + f.Extension);
                    document[DocumentPart.Folder] = folders[1];
                }
                return document;
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
                return document;
            }

        }
        /// <summary>
        /// Creates document information without specifying folders, returning the document with default values.
        /// </summary>
        /// <returns></returns>
        public override Document CreateDocumentInfos()
        {
            return CreateDocumentInfos(null);
        }
        /// <summary>
        /// Saves the document data to the database using the base class method.
        /// </summary>
        public void SaveDocumentData()
        {
            base.SaveDocumentData(document);
        }
        public void Collect()
        {
            base.Collect(document);
        }
        /// <summary>
        /// Gets the document associated with this MeasureFirstPartInfo instance.
        /// </summary>
        /// <returns></returns>
        public override Document GetDocument()
        {
            return document;
        }
    }
    /// <summary>
    /// Represents the information of a VMPB document, including its type, root path, template, and other relevant parts.
    /// </summary>
    public class VmpbDocumentInfo : DocumentInfo
    {
        private Document document;
        ILogger logger;
        public VmpbDocumentInfo(IContainerExtension container) : base(container)
        {
            var factory = container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<VmpbDocumentInfo>();
        }
        /// <summary>
        /// Creates document information based on the provided folders and whether it is a dummy document, returning the document with populated values.
        /// </summary>
        /// <param name="folders"></param>
        /// <param name="isDummy"></param>
        /// <returns></returns>
        public override Document CreateDocumentInfos(string[]? folders, bool isDummy = false)
        {
            try
            {
                document = new VmpbDocument();
                document[DocumentPart.Type] = "VmpbPart";
                document[DocumentPart.RootPath] = string.Empty;
                document[DocumentPart.Template] = string.Empty;
                document[DocumentPart.MaterialRegularEx] = string.Empty;
                document[DocumentPart.JumpTarget] = string.Empty;
                document[DocumentPart.Folder] = string.Empty;
                document[DocumentPart.Template_Size1] = string.Empty;
                document[DocumentPart.Template_Size2] = string.Empty;
                document[DocumentPart.Template_Size3] = string.Empty;
                document[DocumentPart.Template_Size4] = string.Empty;
                document[DocumentPart.OriginalFolder] = string.Empty;
                document[DocumentPart.SavePath] = string.Empty;
                document[DocumentPart.VNR] = string.Empty;
                if (RuleInfo.Rules.Keys.Contains(document[DocumentPart.Type]) == false) return document;
                var xml = XmlSerializerHelper.GetSerializer(typeof(List<Entry>));

                TextReader reader = new StringReader(RuleInfo.Rules[document[DocumentPart.Type]].RuleData);
                List<Entry> doc = (List<Entry>)xml.Deserialize(reader);
                foreach (var entry in doc)
                {
                    DocumentPart DokuPart = (DocumentPart)Enum.Parse(typeof(DocumentPart), entry.Key.ToString());
                    document[DokuPart] = (string)entry.Value;
                }
                if (folders != null)
                {

                    document[DocumentPart.TTNR] = folders[0];
                    if (folders.Length >= 3)
                    {
                        document[DocumentPart.VNR] = folders[2];
                        document[DocumentPart.File] = string.Format("{0}_{1}_{2}.docx", folders[1], folders[2], folders[0]);
                    }
                    Regex regex = new Regex(document[DocumentPart.MaterialRegularEx]);
                    Match match2 = regex.Match(folders[0]);
                    StringBuilder nsb = new StringBuilder();
                    foreach (Group ma in match2.Groups.Values.Skip(1))
                    {
                        if (ma.Value != folders[0])
                        {
                            nsb.Append(ma.Value).Append(Path.DirectorySeparatorChar);
                        }
                    }
                    document[DocumentPart.SavePath] = nsb.Append(folders[1]).Append(Path.DirectorySeparatorChar).ToString();
                    
                }

                return document;
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
                return document;
            }
        }

        public override Document CreateDocumentInfos()
        {
            return CreateDocumentInfos(null);
        }
        public void SaveDocumentData()
        {
            base.SaveDocumentData(document);
        }
        public void Collect()
        {
            base.Collect(document);
        }

        public override Document GetDocument()
        {
            return document;
        }
    }
    /// <summary>
    /// Represents the information of a work area document, including its type, root path, template, and other relevant parts.
    /// </summary>
    public class WorkareaDocumentInfo : DocumentInfo
    {
        private Document document;
        ILogger logger;
        public WorkareaDocumentInfo(IContainerExtension container) : base(container)
        {
            var factory = container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<WorkareaDocumentInfo>();
        }

        public override Document CreateDocumentInfos(string[]? folders, bool isDummy = false)
        {
            try
            {
                document = new FirstPartDocument();
                document[DocumentPart.Type] = "WorkPart";
                document[DocumentPart.RootPath] = string.Empty;
                document[DocumentPart.MaterialRegularEx] = string.Empty;
                document[DocumentPart.DummyRegualarEx] = string.Empty;
                document[DocumentPart.SavePath] = string.Empty;
                if (RuleInfo.Rules.Keys.Contains(document[DocumentPart.Type]) == false) return document;
                var xml = XmlSerializerHelper.GetSerializer(typeof(List<Entry>));

                TextReader reader = new StringReader(RuleInfo.Rules[document[DocumentPart.Type]].RuleData);
                List<Entry> doc = (List<Entry>)xml.Deserialize(reader);
                RegexOptions options = RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace;
                foreach (var entry in doc)
                {
                    DocumentPart DokuPart = (DocumentPart)Enum.Parse(typeof(DocumentPart), entry.Key.ToString().Trim());
                    document[DokuPart] = (string)entry.Value;
                }

                if (folders != null)
                {

                    document[DocumentPart.TTNR] = folders[0];
                    Regex regex;
                    if (isDummy)
                        regex = new(document[DocumentPart.DummyRegualarEx]);
                    else
                        regex = new(document[DocumentPart.MaterialRegularEx]);

                    regex.Options.CompareTo(RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
                    Match match2 = regex.Match(folders[0]);
                    if (!match2.Success) { return document; }
                    StringBuilder nsb = new();
                    foreach (Group ma in match2.Groups.Values.Skip(1))
                    {
                        if (ma.Value != folders[0])
                        {
                            if (!string.IsNullOrEmpty(ma.Value)) nsb.Append(ma.Value).Append(Path.DirectorySeparatorChar);
                        }
                    }
                    foreach (var s in folders.Skip(1))
                    {
                        if (!string.IsNullOrEmpty(s)) nsb.Append(s).Append(Path.DirectorySeparatorChar);
                    }
                    document[DocumentPart.SavePath] = nsb.ToString();
                }

                return document;
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
                return document;
            }
        }

        public override Document CreateDocumentInfos()
        {
            return CreateDocumentInfos(null);
        }
        public void SaveDocumentData()
        {
            base.SaveDocumentData(document);
        }
        public void Collect()
        {
            base.Collect(document);
        }

        public override Document GetDocument()
        {
            return document;
        }
    }
    /// <summary>
    /// Represents the information of a measure document, including its type, root path, template, and other relevant parts.
    /// </summary>
    public class MeasureDocumentInfo : DocumentInfo
    {
        private Document document;
        ILogger logger;
        public MeasureDocumentInfo(IContainerExtension container) : base(container)
        {
            var factory = container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<MeasureDocumentInfo>();
        }

        public override Document CreateDocumentInfos(string[]? folders, bool isDummy = false)
        {
            try
            {
                document = new FirstPartDocument();
                document[DocumentPart.Type] = "MeasurePart";
                document[DocumentPart.RootPath] = string.Empty;
                document[DocumentPart.MaterialRegularEx] = string.Empty;
                document[DocumentPart.JumpTarget] = string.Empty;
                document[DocumentPart.Folder] = string.Empty;

                if (RuleInfo.Rules.Keys.Contains(document[DocumentPart.Type]) == false) return document;
                var xml = XmlSerializerHelper.GetSerializer(typeof(List<Entry>));

                TextReader reader = new StringReader(RuleInfo.Rules[document[DocumentPart.Type]].RuleData);
                List<Entry> doc = (List<Entry>)xml.Deserialize(reader);
                foreach (var entry in doc)
                {
                    DocumentPart DokuPart = (DocumentPart)Enum.Parse(typeof(DocumentPart), entry.Key.ToString());
                    document[DokuPart] = (string)entry.Value;
                }

                if (folders != null)
                {

                    document[DocumentPart.TTNR] = folders[0];
                    Regex regex = new Regex(document[DocumentPart.MaterialRegularEx]);
                    Match match2 = regex.Match(folders[0]);
                    StringBuilder nsb = new StringBuilder();
                    foreach (Group ma in match2.Groups.Values.Skip(1))
                    {
                        if (ma.Value != folders[0])
                        {
                            nsb.Append(ma.Value).Append(Path.DirectorySeparatorChar);
                        }
                    }
                    document[DocumentPart.SavePath] = nsb.Append(folders[1]).Append(Path.DirectorySeparatorChar)
                        .Append(document[DocumentPart.Folder]).Append(Path.DirectorySeparatorChar).ToString();
                }

                return document;
            }
            catch (Exception e)
            {
                logger.LogError("{message}", e.ToString());
                return document;
            }
        }

        public override Document CreateDocumentInfos()
        {
            return CreateDocumentInfos(null);
        }
        public void SaveDocumentData()
        {
            base.SaveDocumentData(document);
        }
        public void Collect()
        {
            base.Collect(document);
        }

        public override Document GetDocument()
        {
            return document;
        }
    }


    public class Entry
    {
        public object Key;
        public object Value;
        public Entry(object key, object value)
        {
            this.Key = key;
            this.Value = value;
        }

        public Entry() { this.Key = new(); this.Value = new(); }
    }
    public enum DocumentPart
    {
        Type,
        RootPath,
        Template,
        MaterialRegularEx,
        Template_Size1,
        Template_Size2,
        Template_Size3,
        SavePath,
        TTNR,
        File,
        JumpTarget,
        Folder,
        RasterFolder1,
        RasterFolder2,
        RasterFolder3,
        OriginalFolder,
        VNR,
        DummyRegualarEx,
        Template_Size4
    }
    public enum DocumentType
    {
        MeasureFirstPart,
        MeasureVMPB
    }


    ///// <summary>
    ///// Generic tree node class
    ///// </summary>
    ///// <typeparam name="T">Node type</typeparam>
    //public class CompositeNode<T> where T : IComparable<T>
    //{
    //    // Add a child tree node
    //    public CompositeNode<T> Add(T child)
    //    {
    //        var newNode = new CompositeNode<T> { Node = child };
    //        Children.Add(newNode);
    //        return newNode;
    //    }
    //    // Remove a child tree node
    //    public void Remove(T child)
    //    {
    //        foreach (var compositeNode in Children)
    //        {
    //            if (compositeNode.Node.CompareTo(child) == 0)
    //            {
    //                Children.Remove(compositeNode);
    //                return;
    //            }
    //        }
    //    }
    //    // Gets or sets the node
    //    public T Node { get; set; } = default!;
    //    // Gets treenode children
    //    public List<CompositeNode<T>> Children { get; } = [];
    //    // Recursively displays node and its children 
    //    public static void Display(CompositeNode<T> node, int indentation)
    //    {
    //        var line = new string('-', indentation);
    //        //WriteLine(line + " " + node.Node);
    //        node.Children.ForEach(n => Display(n, indentation + 1));
    //    }
    //}
    ///// <summary>
    ///// Shape class
    ///// <remarks>
    ///// Implements generic IComparable interface
    ///// </remarks>
    ///// </summary>
    //public class Shape(string name) : IComparable<Shape>
    //{
    //    public override string ToString() => name;
    //    // IComparable<Shape> Member
    //    public int CompareTo(Shape? other) => (this == other) ? 0 : -1;
    //}

}
