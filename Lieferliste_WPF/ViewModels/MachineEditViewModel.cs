using El2Core.Models;
using El2Core.Services;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace Lieferliste_WPF.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public class MachineEditViewModel : ViewModelBase, IViewModel
    {

        public string Title { get; } = "Maschinen Zuteilung";
        public ICommand SaveCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }
        public ICommand TextSearchCommand => textSearchCommand ??= new RelayCommand(OnTextSearch);

        public ICollectionView RessourcesCV { get; }
        ILogger logger;
        private RelayCommand? textSearchCommand;
        private string _searchFilterText = String.Empty;
        private readonly IContainerExtension _container;
        private readonly IUserSettingsService _settingsService;
        private readonly DB_COS_LIEFERLISTE_SQLContext DBctx;
        private ObservableCollection<Ressource>? _ressources { get; set; }
        public static List<WorkArea> WorkAreas { get; } = new();

        public bool HasChange => DBctx.ChangeTracker.HasChanges();

        public MachineEditViewModel(IContainerExtension container, IUserSettingsService settingsService)
        {
            _container = container;
            DBctx = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            _settingsService = settingsService;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<MachineEditViewModel>();
            LoadData();

            RessourcesCV = new ListCollectionView(_ressources);

            SaveCommand = new ActionCommand(OnSaveExecuted, OnSaveCanExecute);
            CloseCommand = new ActionCommand(OnCloseExecuted, OnCloseCanExecute);

            RessourcesCV.MoveCurrentToFirst();
            RessourcesCV.CurrentChanged += OnChanged;
            RessourcesCV.Filter += RessView_FilterPredicate;

        }


        private bool RessView_FilterPredicate(object value)
        {
            var val = value as Ressource;
            Ressource res = (Ressource)value;

            bool accepted = true;
            int v;
            if (!string.IsNullOrWhiteSpace(_searchFilterText))
            {
                _searchFilterText = _searchFilterText.ToUpper();
                if (int.TryParse(_searchFilterText, out v))
                {
                    if (!(accepted = res.Inventarnummer?.ToUpper().StartsWith(_searchFilterText) ?? false))
                        accepted = res.RessourceCostUnits.Any(y => y.CostId.Equals(v));
                }
                else
                {
                    if (!(accepted = (res.RessName != null) && res.RessName.ToUpper().Contains(_searchFilterText)))
                        accepted = (res.WorkArea?.Bereich != null) && res.WorkArea.Bereich.ToUpper().Contains(_searchFilterText);
                }
            }
            return accepted;
        }
        private void OnTextSearch(object commandParameter)
        {
            if (commandParameter is TextChangedEventArgs change)
            {
                TextBox tb = (TextBox)change.OriginalSource;
                if (tb.Text.Length >= 3 || tb.Text.Length == 0)
                {
                    _searchFilterText = tb.Text;

                    var uiContext = SynchronizationContext.Current;
                    uiContext?.Send(x => RessourcesCV.Refresh(), null);
                }
            }
        }
        private void OnChanged(object? sender, EventArgs e)
        {


        }

        private void OnCloseExecuted(object obj)
        {
            if (obj is Window ob)
            {
                //if (DBctx.ChangeTracker.HasChanges())
                //{
                //    MessageBoxResult result = MessageBox.Show("Wollen Sie die Änderungen noch Speichern?", "Datenbank Speichern"
                //        , MessageBoxButton.YesNo, MessageBoxImage.Question);
                //    if (result == MessageBoxResult.Yes) { DBctx.SaveChangesAsync(); }

                //}

                ob.Close();

            }
        }

        private bool OnCloseCanExecute(object arg)
        {
            return true;
        }

        private void OnSaveExecuted(object obj)
        {
            DBctx.SaveChanges();

        }

        private bool OnSaveCanExecute(object arg)
        {
            return true;
        }

        private void LoadData()
        {
            _ressources = DBctx.Ressources
                .Include(y => y.RessourceCostUnits)
                .Include(x => x.WorkArea)
                .Where(y => y.Inventarnummer != null)
                .ToObservableCollection();

            var w = DBctx.WorkAreas.AsNoTracking()
                .OrderBy(x => x.Bereich)
                .Select(y => new WorkArea() { WorkAreaId = y.WorkAreaId, Bereich = y.Bereich })
                .ToList();
            WorkAreas.AddRange(w.Prepend(new WorkArea() { WorkAreaId = 0, Bereich = "nicht zugeteilt" }));

        }

        public void Closing()
        {
            try
            {
                if (DBctx.ChangeTracker.HasChanges())
                {
                    if (_settingsService.IsSaveMessage)
                    {
                        var result = MessageBox.Show(string.Format("Sollen die Änderungen in {0} gespeichert werden?", Title),
                            Title, MessageBoxButton.YesNo, MessageBoxImage.Question);
                        if (result == MessageBoxResult.Yes)
                        {
                            DBctx.SaveChanges();
                        }
                    }
                    else DBctx.SaveChanges();
                }
            }
            catch (Exception e)
            {

                logger.LogError("{message}", e.ToString());
            }
        }
    }
    public class MachineNameValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (value is string val)
            {
                if (string.IsNullOrEmpty(val)) { return new ValidationResult(false, "Der Eintrag darf nicht leer sein"); }
            }
            return ValidationResult.ValidResult;
        }
    }
}
