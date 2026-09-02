using El2Core.Models;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Xml.Serialization;

namespace El2Core.Utils
{
    public interface IPersonalFilterContainer
    {
        IContainerProvider Container { get; }
    }
    /// <summary>
    /// Singleton class that manages a collection of personal filters.
    /// It provides methods to add, remove, load, save, and access filters by key.
    /// The filters are serialized to and deserialized from an XML file named "Perfilter.xml" located in the user's roaming and local configuration directory.
    /// </summary>
    public sealed class PersonalFilterContainer
    {
        private readonly Dictionary<string, PersonalFilter> _filters = [];
        private static readonly PersonalFilterContainer Instance = new ();

        private PersonalFilterContainer()
        {
            _filters.Add("_keine", null);
            Load ();
        }

        public PersonalFilter this[string key]
        {
            get => _filters[key];
            set => _filters[key] = value;
        }
        public string[] Keys => [.. _filters.Keys];

        /// <summary>
        /// Gets the singleton instance of the PersonalFilterContainer.
        /// </summary>
        /// <returns></returns>
        public static PersonalFilterContainer GetInstance()
        {
            return Instance;
        }
        public bool IsChanged { get; private set; } = false;


        private void Load()
        {
            var Configfile = new FileInfo(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath);
            var Folder = Configfile.Directory.Parent.Parent.FullName;
            
            FileInfo fileInfo = new FileInfo(Path.Combine(Folder, "Perfilter.xml"));

            if (fileInfo.Exists)
            {
                try
                {
                    DeserializeObject(fileInfo.FullName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

        }
        public void Reload()
        {
            _filters.Clear();
            
            Load();
        }
        public void Remove(string key)
        {
            _filters.Remove(key);
            IsChanged = true;
        }
        public void Add(string name, PersonalFilter filter)
        {
            _filters.Add(name, filter);
            filter.PropertyChanged += OnFilterPropertyChanged;
            IsChanged = true;
        }

        private void OnFilterPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            IsChanged = true;
        }
        /// <summary>
        /// Saves the current state of the personal filters to an XML file named "Perfilter.xml" in the user's roaming and local configuration directory.
        /// </summary>
        public void Save()
        {
            if (IsChanged)
            {

                var Configfile = new FileInfo(ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath);
                var Folder = Configfile.Directory.Parent.Parent.FullName;
                string fileName = Path.Combine(Folder, "Perfilter.xml");

                SerializeObject(fileName);
                IsChanged = false;
                Reload();
            }
        }

        public void SerializeObject(string filename)
        {
            try
            {
                // Each overridden field, property, or type requires
                // an XmlAttributes instance.  
                XmlAttributes attrs = new XmlAttributes();
                XmlElementAttribute attr = new XmlElementAttribute();
                attr.ElementName = "Filter";
                attr.Type = typeof(PersonalFilterVorgang);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                attr.Type = typeof(PersonalFilterOrderRb);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                attr.Type = typeof(PersonalFilterMaterial);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                attr.Type = typeof(PersonalFilterRessource);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                attr.Type = typeof(PersonalFilterProject);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                // Creates the XmlAttributeOverrides instance.  
                XmlAttributeOverrides attrOverrides = new XmlAttributeOverrides();

                // Adds the type of the class that contains the overridden
                // member, as well as the XmlAttributes instance to override it
                // with, to the XmlAttributeOverrides.  
                attrOverrides.Add(typeof(PersonalFilter), "Filters", attrs);

                // Creates the XmlSerializer using the XmlAttributeOverrides.  
                XmlSerializer s =
                new XmlSerializer(typeof(List<PersonalFilter>), attrOverrides);

                // Writing the file requires a TextWriter instance.
                var f = new FileInfo(filename);
                TextWriter writer = new StreamWriter(filename);

                // Creates the object to be serialized.  
                List<PersonalFilter> filters = new List<PersonalFilter>();
                foreach (var filter in _filters)
                {
                    if(filter.Value != null)
                        filters.Add(filter.Value);
                }

                // Serializes the object.  
                s.Serialize(writer, filters);
                writer.Close();
            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message, "Serialize", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void DeserializeObject(string filename)
        {
            try
            {
                XmlAttributeOverrides attrOverrides =
              new XmlAttributeOverrides();
                XmlAttributes attrs = new XmlAttributes();

                XmlElementAttribute attr = new XmlElementAttribute();
                attr.ElementName = "Filter";
                attr.Type = typeof(PersonalFilterVorgang);

                // Adds the XmlElementAttribute to the collection of objects.  
                attrs.XmlElements.Add(attr);
                attr.Type = typeof(PersonalFilterOrderRb);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);

                attr.Type = typeof(PersonalFilterProject);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);

                attr.Type = typeof(PersonalFilterMaterial);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);

                attr.Type = typeof(PersonalFilterRessource);

                // Adds the element to the collection of elements.  
                attrs.XmlElements.Add(attr);
                attrOverrides.Add(typeof(PersonalFilter), "Filters", attrs);

                // Creates the XmlSerializer using the XmlAttributeOverrides.  
                XmlSerializer s =
                new XmlSerializer(typeof(PersonalFilter[]), attrOverrides);

                FileStream fs = new FileStream(filename, FileMode.Open);
                var filters = (PersonalFilter[])s.Deserialize(fs);
                _filters.TryAdd("_keine", null);
                foreach (var filter in filters)
                {
                    _filters.Add(filter.Name, filter);
                    filter.PropertyChanged += OnFilterPropertyChanged;                   
                }
            }
            catch (Exception e)
            {

                MessageBox.Show(e.ToString(), "Deserialize", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    [XmlInclude(typeof(PersonalFilterVorgang))]
    [XmlInclude(typeof (PersonalFilterOrderRb))]
    [XmlInclude(typeof(PersonalFilterMaterial))]
    [XmlInclude(typeof(PersonalFilterRessource))]
    [XmlInclude(typeof(PersonalFilterProject))]
    [Serializable]
    public abstract partial class PersonalFilter : INotifyPropertyChanged
    {
        public abstract string Name {get; set;}
        public abstract string Pattern { get; set; }
        public abstract (string, string, int) Field { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;

        public abstract Regex GetRegEx();

        public abstract string GetTestString(Vorgang vorgang, IContainerProvider container);


        public bool TestValue(Vorgang vorgang, IContainerProvider container)
        {
            var Reg = GetRegEx();
            var test = GetTestString(vorgang, container);
            return (Reg != null) ? Reg.Match(test).Success : false;
        }
        protected void OnPropertyChanged(string propertyName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = "")
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
        [GeneratedRegex("")]
        public static partial Regex MyRegex();
    }
 
    [Serializable]
    public partial class PersonalFilterVorgang : PersonalFilter
    {
        public PersonalFilterVorgang() { }
        public PersonalFilterVorgang(string name, string regex, (string, string, int) field) { _name = name; RegEx = new Regex(regex); _Field = field; }
        private Regex RegEx = MyRegex();
        private string _name = string.Empty;
        public override string Name { get => _name; set => SetField(ref _name, value); }
        private (string, string, int) _Field;
        public override (string, string, int) Field
        {
            get => _Field;
            set => SetField(ref _Field, value);
        }
        public override string? Pattern
        {
            get { return (RegEx != null) ? RegEx.ToString() : null; }
            set
            {
                RegEx = new Regex(value);
                OnPropertyChanged(nameof(Pattern));
            }
        }

   
        public override Regex GetRegEx()
        {
            return RegEx;
        }

        public override string GetTestString(Vorgang vorgang, IContainerProvider container)
        {
            PropertyInfo? info = vorgang.GetType().GetProperty(Field.Item2);
            if (info != null)
                return info.GetValue(vorgang, null)?.ToString() ?? string.Empty;
            return string.Empty;
        }


    }
    [Serializable]
    public class PersonalFilterOrderRb : PersonalFilter
    {
        private string _name = string.Empty;

        public override string Name { get => _name; set => SetField(ref _name, value); }
        private Regex RegEx = MyRegex();
        private (string, string, int) _Field;
 
        public PersonalFilterOrderRb() { }
        public PersonalFilterOrderRb(string name, string regex, (string, string, int) field)
        {
            _name = name;
            RegEx = new Regex(regex);
            _Field = field;
        }
        /// <summary>
        /// Gets or sets the regular expression pattern associated with this personal filter.
        /// When set, it updates the RegEx property and raises the PropertyChanged event for the Pattern property.
        /// </summary>
        public override string Pattern
        {
            get { return RegEx.ToString(); }
            set
            {
                RegEx = new Regex(value);
                OnPropertyChanged(nameof(Pattern));
            }
        }

        public override (string, string, int) Field
        {
            get => _Field;
            set => SetField(ref _Field, value);
        }
        /// <summary>
        /// Gets the regular expression associated with this personal filter.
        /// </summary>
        /// <returns></returns>
        public override Regex GetRegEx()
        {
            return RegEx;
        }
        /// <summary>
        /// Gets the test string for the specified Vorgang by navigating through the AidNavigation property to retrieve the value of the property specified in Field.Item2.
        /// </summary>
        /// <param name="vorgang"></param>
        /// <param name="container"></param>
        /// <returns></returns>
        public override string GetTestString(Vorgang vorgang, IContainerProvider container)
        {

            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>(); 
 
            var modelData = db.Vorgangs.EntityType;

            var nav = modelData.FindDeclaredNavigation("AidNavigation");
            if (nav != null)
            {
                modelData = nav.TargetEntityType;
                var value = modelData?.FindDeclaredProperty(Field.Item2)?.PropertyInfo?.GetValue(vorgang.AidNavigation, null);
                return (value != null) ? value.ToString() : string.Empty;
            }
                     
            return string.Empty;
        }
    }
    [Serializable]
    public class PersonalFilterMaterial : PersonalFilter
    {
        private string _name = string.Empty; 
        public override string Name { get => _name; set => SetField(ref _name, value); }
        private Regex RegEx = MyRegex();
        private (string, string, int) _Field;
        private PersonalFilterMaterial() { }
        public PersonalFilterMaterial(string name, string regex, (string, string, int) field)
        {
            _name = name;
            RegEx = new Regex(regex);
            _Field = field;
        }
        public override string Pattern
        {
            get { return RegEx.ToString(); }
            set
            {
                RegEx = new Regex(value);
                OnPropertyChanged(nameof(Pattern));
            }
        }

        public override (string, string, int) Field
        {
            get => _Field;
            set => SetField(ref _Field, value);
        }

        public override Regex GetRegEx()
        {
            return RegEx;
        }
        /// <summary>
        /// Gets the test string for the specified Vorgang by navigating through the AidNavigation
        /// and MaterialNavigation properties to retrieve the value of the property specified in Field.Item2.
        /// </summary>
        /// <param name="vorgang"></param>
        /// <param name="container"></param>
        /// <returns></returns>
        public override string GetTestString(Vorgang vorgang, IContainerProvider container)
        {

            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            var modelData = db.Vorgangs.EntityType;

            var nav = modelData.FindDeclaredNavigation("AidNavigation");
            modelData = nav.TargetEntityType;
            nav = modelData.FindDeclaredNavigation("MaterialNavigation");
            if (nav != null)
            {

                modelData = nav.TargetEntityType;
                if (vorgang.AidNavigation.MaterialNavigation != null)
                {
                    var value = modelData?.FindDeclaredProperty(Field.Item2)?.PropertyInfo?.GetValue(vorgang.AidNavigation.MaterialNavigation, null);
                    return (value != null) ? value.ToString() : string.Empty;
                }
            }

            return string.Empty;
        }
    }
    [Serializable]
    public class PersonalFilterRessource : PersonalFilter
    {
        private string _name = string.Empty;
        public override string Name { get => _name; set => SetField(ref _name, value); }
        private Regex RegEx = MyRegex();
        private (string, string, int) _Field;
        public PersonalFilterRessource() { }
        public PersonalFilterRessource(string name, string regex, (string, string, int) field)
        {
            _name = name;
            RegEx = new Regex(regex);
            _Field = field;
        }
        public override string Pattern
        {
            get { return RegEx.ToString(); }
            set
            {
                RegEx = new Regex(value);
                OnPropertyChanged(nameof(Pattern));
            }
        }

        public override (string, string, int) Field
        {
            get => _Field;
            set => SetField(ref _Field, value);
        }

        public override Regex GetRegEx()
        {
            return RegEx;
        }
        /// <summary>
        /// Gets the test string for the specified Vorgang by navigating through the RidNavigation property to retrieve the value of the property specified in Field.Item2.
        /// </summary>
        /// <param name="vorgang"></param>
        /// <param name="container"></param>
        /// <returns></returns>
        public override string GetTestString(Vorgang vorgang, IContainerProvider container)
        {

            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            var modelData = db.Vorgangs.EntityType;

            var nav = modelData.FindDeclaredNavigation("RidNavigation");
            if (nav != null)
            {
                modelData = nav.TargetEntityType;
                if (vorgang.RidNavigation != null)
                {
                    var value = modelData?.FindDeclaredProperty(Field.Item2)?.PropertyInfo?.GetValue(vorgang.RidNavigation, null);
                    return (value != null) ? value.ToString() : string.Empty;
                }
            }

            return string.Empty;
        }
    }
    [Serializable]
    public class PersonalFilterProject : PersonalFilter
    {
        private string _name = string.Empty;
        public override string Name { get => _name; set => SetField(ref _name, value); }
        private Regex RegEx = MyRegex();
        private (string, string, int) _Field;
        public PersonalFilterProject() { }
        public PersonalFilterProject(string name, string regex, (string, string, int) field)
        {
            _name = name;
            RegEx = new Regex(regex);
            _Field = field;
        }
        public override string Pattern
        {
            get { return RegEx.ToString(); }
            set
            {
                RegEx = new Regex(value);
                OnPropertyChanged(nameof(Pattern));
            }
        }

        public override (string, string, int) Field
        {
            get => _Field;
            set => SetField(ref _Field, value);
        }

        public override Regex GetRegEx()
        {
            return RegEx;
        }
        /// <summary>
        /// Gets the test string for the specified Vorgang by navigating through the AidNavigation and Pro properties to retrieve the value of the property specified in Field.Item2.
        /// </summary>
        /// <param name="vorgang"></param>
        /// <param name="container"></param>
        /// <returns></returns>
        public override string GetTestString(Vorgang vorgang, IContainerProvider container)
        {

            using var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            var modelData = db.Vorgangs.EntityType;

            var nav = modelData.FindDeclaredNavigation("AidNavigation");
            modelData = nav?.TargetEntityType;
            nav = modelData?.FindDeclaredNavigation("Pro");
            if (nav != null)
            {
                modelData = nav.TargetEntityType;
                if (vorgang.AidNavigation.Pro != null)
                {
                    var value = modelData?.FindDeclaredProperty(Field.Item2)?.PropertyInfo?.GetValue(vorgang.AidNavigation.Pro, null);
                    return (value != null) ? value.ToString() : string.Empty;
                }
                return string.Empty;
            }

            return string.Empty;
        }
    }
    /// <summary>
    /// Represents a pair of properties with associated metadata.
    /// </summary>
    public readonly struct PropertyPair
    {
        //Type 1 == Vorgang
        //Type 2 == OrderRb
        //Type 3 == Material
        //Type 4 == Ressource
        //Type 5 == Project
        public static ValueTuple<string, string, int> OrderNumber = ValueTuple.Create("Auftragsnummer", "Aid", 1);
        public static ValueTuple<string, string, int> ProcessDescription = ValueTuple.Create("KurzText", "Text", 1);
        public static ValueTuple<string, string, int> Material = ValueTuple.Create("Material", "Material", 2);
        public static ValueTuple<string, string, int> MaterialDescription = ValueTuple.Create("MaterialBezeichnung", "Bezeichng", 3);
        public static ValueTuple<string, string, int> RessourceName = ValueTuple.Create("Maschinenname", "RessName", 4);
        public static ValueTuple<string, string, int> LieferTermin = ValueTuple.Create("Liefertermin", "Liefertermin", 2);
        public static ValueTuple<string, string, int> PrioText = ValueTuple.Create("PrioText", "Prio", 2);
        public static ValueTuple<string, string, int> Project = ValueTuple.Create("Projekt", "ProId", 2);
        public static ValueTuple<string, string, int> ProjectInfo = ValueTuple.Create("Projekt Info", "ProjectInfo", 5);
        public static ValueTuple<string, string, int> MarkerCode = ValueTuple.Create("Reihung", "MarkCode", 2);
        
    }

}
