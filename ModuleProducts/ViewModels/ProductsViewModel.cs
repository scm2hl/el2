using CompositeCommands.Core;
using El2Core.Constants;
using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace ModuleProducts.ViewModels
{
    internal class ProductsViewModel : ViewModelBase
    {
        public ProductsViewModel(IContainerExtension container, IApplicationCommands applicationCommands, IDialogService dialogService)
        {
            _container = container;
            _applicationCommands = applicationCommands;
            var loggerFactory = _container.Resolve<ILoggerFactory>();
            _Logger = loggerFactory.CreateLogger<ProductsViewModel>();
            firstPartInfo = new MeasureFirstPartInfo(_container);
            MaterialTask = new NotifyTaskCompletion<ICollectionView>(OnLoadMaterialsAsync());
            _dialogService = dialogService;
        }
        public string Title { get; } = "Produkt Übersicht";

        private readonly IContainerExtension _container;
        private readonly ILogger _Logger;
        private readonly IDialogService _dialogService;
        public ICollectionView ProductsView { get; private set; }
        private MeasureFirstPartInfo firstPartInfo;
        private ObservableCollection<ProductMaterial> _Materials =[];
        private int _ArchivProcessingCount;

        public int ArchivProcessingCount
        {
            get { return _ArchivProcessingCount; }
            set
            {
                _ArchivProcessingCount = value;
                NotifyPropertyChanged(() => ArchivProcessingCount);
            }
        }
        private int _Archivated = 0;

        public int Archivated
        {
            get { return _Archivated; }
            set
            {
                _Archivated = value;
                NotifyPropertyChanged(() => Archivated);
            }
        }
        private int _ArchivState2Count = 0;

        public int ArchivState2Count
        {
            get { return _ArchivState2Count; }
            set
            {
                _ArchivState2Count = value;
                NotifyPropertyChanged(() => ArchivState2Count);
            }
        }
        private int _ArchivState3Count = 0;

        public int ArchivState3Count

        {
            get { return _ArchivState3Count; }
            set
            {
                _ArchivState3Count = value;
                NotifyPropertyChanged(() => ArchivState3Count);
            }
        }
        private int _ArchivState4Count = 0;

        public int ArchivState4Count
        {
            get { return _ArchivState4Count; }
            set
            {
                _ArchivState4Count = value;
                NotifyPropertyChanged(() => ArchivState4Count);
            }
        }
        private int _MovedFiles;

        public int MovedFiles
        {
            get { return _MovedFiles; }
            set
            {
                _MovedFiles = value;
                NotifyPropertyChanged(() => MovedFiles);
            }
        }

        private bool _ArchivComplete = false;

        public bool ArchivComplete
        {
            get { return _ArchivComplete; }
            set
            {
                _ArchivComplete = value;
                NotifyPropertyChanged(() => ArchivComplete);
            }
        }
        private bool _IsArchivating;

        public bool IsArchivating
        {
            get { return _IsArchivating; }
            set
            {
                _IsArchivating = value;
                NotifyPropertyChanged(() => IsArchivating);
            }
        }
        private bool? _IsSearchMSF = null;
        public bool? IsSearchMSF
        {
            get { return _IsSearchMSF; }
            set
            {
                ProductsCount = 0;
                _IsSearchMSF = value;
                ProductsView.Refresh();
            }
        }
        private bool? _IsSearchClosed = null;
        public bool? IsSearchClosed
        {
            get { return _IsSearchClosed; }
            set
            {
                ProductsCount = 0;
                _IsSearchClosed = value;
                ProductsView.Refresh();
            }
        }
        private int _ProductsCount;

        public int ProductsCount
        {
            get { return _ProductsCount; }
            set
            {
                _ProductsCount = value;
                NotifyPropertyChanged(() => ProductsCount);
            }
        }

        private string? _SearchText;
        public string? SearchText
        {
            get { return _SearchText; }
            set
            {
                ProductsCount = 0;
                _SearchText = value;
                OnTextSearch(value);
            }
        }
        private DateTime _StartDateFilter;

        public DateTime StartDateFilter
        {
            get { return _StartDateFilter; }
            set
            {
                ProductsCount = 0;
                _StartDateFilter = value;
                if (value != null && EndDateFilter != null && value >= EndDateFilter)
                    throw new Exception("Startdatum muss kleiner als Enddatum sein");
                
            }
        }
        private IEnumerable<DateTime> _Selected_Dates;

        public IEnumerable<DateTime> Selected_Dates
        {
            get { return _Selected_Dates; }
            set { _Selected_Dates = value; }
        }

        private DateTime? _EndDateFilter;

        public DateTime? EndDateFilter
        {
            get { return _EndDateFilter; }
            set
            {
                _EndDateFilter = value;
                if (value != null && StartDateFilter != null && value < StartDateFilter)
                    throw new Exception("Startdatum muss kleiner als Enddatum sein");
            }
        }

        private RelayCommand? _SearchCommand;
        public RelayCommand SearchCommand => _SearchCommand ??= new RelayCommand(OnTextSearch);
        public ICommand ArchivateCommand => new ActionCommand(OnArchivateExecute, OnCanArchivateExecute);
        public ICommand DateSelectedCommand => new ActionCommand(OnDateSelectedExecute, OnCanDateSelectedExecute);
        public ICommand CloseArchivMessageCommand => new ActionCommand(OnCloseArchivMessageExecuted, OnCanCloseArchivMessageExecute);



        private IApplicationCommands _applicationCommands;
        public IApplicationCommands ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != null)
                {
                    _applicationCommands = value;
                    NotifyPropertyChanged(() => ApplicationCommands);
                }
            }
        }
        private NotifyTaskCompletion<ICollectionView>? _materialTask;

        public NotifyTaskCompletion<ICollectionView>? MaterialTask
        {
            get { return _materialTask; }
            set
            {
                if (_materialTask != value)
                {
                    _materialTask = value;
                    NotifyPropertyChanged(() => MaterialTask);
                }
            }
        }
        private async Task<ICollectionView> OnLoadMaterialsAsync()
        {
            try
            {
                
                using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var materials = await db.GetProcedures().ProductMaterialsAsync(UserInfo.User.UserId);

                _Materials.Clear();
                foreach (var m in materials.GroupBy(x => x.TTNR))
                { 
                    var p = new ProductMaterial(m.Key, m.First().Bezeichng, [.. m]);
                    _Materials.Add(p);               
                }

                ProductsView = new ListCollectionView(_Materials);
                ProductsView.Filter += OnFilterPredicate;
            }
            catch (Exception e)
            {
                _Logger.LogError("OnLoadMaterialsAsync Exception: {Message}\n{StackTrace}", e.Message, e.StackTrace);
            }
            return ProductsView;
        }
        private bool OnCanDateSelectedExecute(object arg)
        {
            return true;
        }

        private void OnDateSelectedExecute(object obj)
        {
            if (obj is DateTime o)
            {
                Selected_Dates = [];
                Selected_Dates = Selected_Dates.Append(o);
            }
            else 
                Selected_Dates = (IEnumerable<DateTime>)obj;
        }
        private bool OnCanCloseArchivMessageExecute(object arg)
        {
            return ArchivComplete;
        }

        private void OnCloseArchivMessageExecuted(object obj)
        {
            ArchivProcessingCount = 0;
            Archivated = 0;
            ArchivState2Count = 0;
            ArchivState3Count = 0;
            ArchivState4Count = 0;
            MovedFiles = 0;
            ArchivComplete = false;
            IsArchivating = false;
        }
        private bool OnCanArchivateExecute(object arg)
        {
            
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.Archivate) && !IsArchivating;
        }
        private void OnArchivateExecute(object obj)
        {
            IsArchivating = true;
            ArchivProcessingCount = 0;
            Archivated = 0;
            ArchivState2Count = 0;
            ArchivState3Count = 0;
            ArchivState4Count = 0;
            MovedFiles = 0;
            ArchivComplete = OnArchivateExecuteAsync(obj).IsCompleted;
        }
        private async Task OnArchivateExecuteAsync(object obj)
        {
            int maxConcurrentTasks = 5; 
            var semaphore = new SemaphoreSlim(maxConcurrentTasks);
            var isArchivated = new ConcurrentBag<(string, string, int, string)>();
            var tasks = new List<Task>();

            // Gather all eligible orders first
            var eligibleOrders = new List<(ProductMaterial mat, ProductOrder ord)>();
            foreach (var m in ProductsView)
            {
                var mat = (ProductMaterial)m;
                foreach (ProductOrder ord in mat.ProdOrders)
                {
                    if (ord.ArchivState == Archivator.ArchivState.None &&
                        ord.Closed &&
                        ord.Completed < DateTime.Now.AddDays(-Archivator.DelayDays))
                    {
                        eligibleOrders.Add((mat, ord));
                    }
                }
            }
            ArchivProcessingCount = eligibleOrders.Count;

            // Shared counters for thread-safe updates
            int archivated = 0, state2 = 0, state3 = 0, state4 = 0, movedFiles = 0;

            foreach (var (mat, ord) in eligibleOrders)
            {
                await semaphore.WaitAsync();
                tasks.Add(Task.Run(async () =>
                {
                    (string, string, int, string)? link = null;
                    try
                    {
                        var doku = firstPartInfo.CreateDocumentInfos([mat.TTNR, ord.OrderNr]);
                        int rulenr = 0;
                        bool matched = false;
                        foreach (var rule in Archivator.ArchiveRules)
                        {
                            string? input = rule.MatchTarget.Equals(Archivator.ArchivatorTarget.TTNR) ? mat.TTNR : ord.OrderNr;
                            if (Regex.IsMatch(input, rule.RegexString))
                            {
                                matched = true;
                                break;
                            }
                            rulenr++;
                        }
                        if (!matched)
                        {
                            Interlocked.Increment(ref state4);
                            return;
                        }

                        var p = Path.Combine(doku[DocumentPart.RootPath], doku[DocumentPart.SavePath], doku[DocumentPart.Folder]);
                        _Logger.LogInformation($"Archivate {p}");

                        var result = await Archivator.ArchivateAsync(new DirectoryInfo(p), rulenr);

                        if (result.State == Archivator.ArchivState.Archivated ||
                            result.State == Archivator.ArchivState.NoFiles)
                        {
                            CoreFunction.DeleteDirectoryWithWait(p, true);
                        }

                        using var db = _container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                        var o = db.OrderRbs.Single(x => x.Aid == ord.OrderNr);

                        switch (result.State)
                        {
                            case Archivator.ArchivState.Archivated:
                                Interlocked.Increment(ref archivated);
                                Interlocked.Add(ref movedFiles, result.MovedFiles);
                                o.ArchivPath = Path.Combine(result.Location, ord.OrderNr);
                                o.ArchivState = (int)result.State;
                                link = (mat.TTNR, ord.OrderNr, (int)result.State, o.ArchivPath);
                                break;
                            case Archivator.ArchivState.NoFiles:
                                Interlocked.Increment(ref state2);
                                o.ArchivState = (int)result.State;
                                break;
                            case Archivator.ArchivState.NoDirectory:
                                Interlocked.Increment(ref state3);
                                o.ArchivState = (int)result.State;
                                break;
                        }
                        if (link != null)
                            isArchivated.Add(link.Value);
                        else
                            isArchivated.Add((mat.TTNR, ord.OrderNr, (int)result.State, string.Empty));

                        db.Update(o);
                        await db.SaveChangesAsync();
                    }
                    catch (Exception ex)
                    {
                        _Logger.LogInformation(ex.Message);
                    }
                    finally
                    {
                        // Decrement ArchivProcessingCount as each task completes
                        Interlocked.Decrement(ref _ArchivProcessingCount);
                        NotifyPropertyChanged(() => ArchivProcessingCount);
                        semaphore.Release();
                    }
                }));
            }

            await Task.WhenAll(tasks);

            // Update UI-bound properties on the UI thread
            Archivated = archivated;
            ArchivState2Count = state2;
            ArchivState3Count = state3;
            ArchivState4Count = state4;
            MovedFiles = movedFiles;
            ArchivProcessingCount = 0;

            // Update ProductOrder states
            foreach (var link in isArchivated)
            {
                var mat = ProductsView.Cast<ProductMaterial>().Single(x => x.TTNR == link.Item1);
                var ord = mat.ProdOrders.Cast<ProductOrder>().Single(x => x.OrderNr == link.Item2);
                ord.ArchivState = (Archivator.ArchivState)link.Item3;
                ord.ArchivPath = link.Item4;
                ord.OrderLink = (link.Item1, link.Item2, link.Item3, link.Item4);
            }
            ProductsView.Refresh();
            _Logger.LogInformation("Archiviert: {0} NoFiles(2): {1} NoDirectory(3): {2} NoRules(4): {3} copied Files {4}",
                Archivated, ArchivState2Count, ArchivState3Count, ArchivState4Count, MovedFiles);
            
        }
 
        private bool OnFilterPredicate(object obj)
        {
            if (obj is not ProductMaterial mat)
                return true;

            // Get original orders collection
            var sourceOrders = mat.ProdOrders.SourceCollection?.Cast<ProductOrder>().ToList() ?? mat.ProdOrders.Cast<ProductOrder>().ToList();

            bool MatLevelSearchMatch()
            {
                return !string.IsNullOrEmpty(_SearchText) &&
                       (mat.TTNR.Contains(_SearchText, StringComparison.CurrentCultureIgnoreCase) ||
                        (mat.Description != null && mat.Description.Contains(_SearchText, StringComparison.CurrentCultureIgnoreCase)));
            }

            bool OrderMatches(ProductOrder ord)
            {
                // Search: if material-level search matched, do not require OrderNr match
                if (!string.IsNullOrEmpty(_SearchText) && !MatLevelSearchMatch())
                {
                    if (!ord.OrderNr.Contains(_SearchText, StringComparison.CurrentCultureIgnoreCase))
                        return false;
                }

                if (Selected_Dates != null && Selected_Dates.Any())
                {
                    if (ord.Completed == null || !Selected_Dates.Contains(ord.Completed.Value.Date))
                        return false;
                }

                if (IsSearchMSF != null)
                {
                    var hasMsf = ord.Tags != null && ord.Tags.Any();
                    if (IsSearchMSF == true && !hasMsf) return false;
                    if (IsSearchMSF == false && hasMsf) return false;
                }

                if (IsSearchClosed != null)
                {
                    if (IsSearchClosed == true && !ord.Closed) return false;
                    if (IsSearchClosed == false && ord.Closed) return false;
                }

                return true;
            }

            // Count matching orders and set the per-material filter
            var matches = sourceOrders.Where(OrderMatches).ToList();
            ProductsCount += matches.Count;
            mat.ProdOrders.Filter = item => item is ProductOrder ord && OrderMatches(ord);

            return matches.Any();
        }
        private void OnTextSearch(object obj)
        {
            if(obj is string search)
            {
                ProductsCount = 0;
                _SearchText = search;
                ProductsView.Refresh();
            }
        }
        public class ProductMaterial
        {
            public string TTNR { get; }
            public string? Description { get; }

            public ICollectionView ProdOrders { get; private set; }
            public ProductMaterial(string ttnr, string? description, List<ProductMaterialsResult> orders)
            {
                TTNR = ttnr;
                Description = description;
                List<ProductOrder> products = [];
                foreach (var order in orders.GroupBy(x => x.AID))
                {   
                    
                    if (order.Any())
                    {
                        var ord = order.First();
                        var d = order.MaxBy(static x => x.VNR)?.Quantityyield;
                        var s = order.Sum(x => x.Quantityscrap);
                        var r = order.Sum(x => x.Quantityrework);
                        var dic = new ValueTuple<string, string, int, string?>(ttnr, order.Key, ord.ArchivState, ord.ArchivPath);
                        var msf = order.Where(x => x.MSF != null).Select(x => x.MSF).ToArray();
                        
                        products.Add(new ProductOrder(dic, order.Key, ord.Quantity, ord.Eckstart, ord.Eckende,
                            d, s, r, ord.abgeschlossen, msf, ord.CompleteDate, (Archivator.ArchivState)ord.ArchivState));
                    }
                }
                ProdOrders = CollectionViewSource.GetDefaultView(products);
            }
            internal class DateValidationRule : ValidationRule
            {
                public override ValidationResult Validate(object value, System.Globalization.CultureInfo cultureInfo)
                {
                    DateTime? date = value as DateTime?;
                    if (false)
                        return new ValidationResult(false, "Datum ist ungültig");
                    return ValidationResult.ValidResult;
                }
            }
        }
        public struct ProductOrder(ValueTuple<string, string, int, string> OrderLink, string OrderNr, int? Quantity,
            DateTime? EckStart, DateTime? EckEnd, int? Delivered, int? Scrap, int? Rework, bool closed,
            string?[] tags, DateTime? completed, Archivator.ArchivState archivState) : INotifyPropertyChanged
        {
            public string OrderNr { get; } = OrderNr;
            public int Quantity { get; } = Quantity ??= 0;
            public bool Closed { get; } = closed;
            private ValueTuple<string, string, int, string> _orderLink = OrderLink;
            public ValueTuple<string, string, int, string> OrderLink
            {
                get { return _orderLink; }
                set {
                    _orderLink = value;
                    OnPropertyChanged(nameof(OrderLink));
                }
            }
            public DateTime? Start { get; } = EckStart;
            public DateTime? End { get; } = EckEnd;
            public int Delivered { get; } = Delivered ??= 0;
            public int Scrap { get; } = Scrap ??= 0;
            public int Rework { get; } = Rework ??= 0;
            public string?[] Tags { get; } = tags;
            public DateTime? Completed { get; } = completed;
            private string? _archivPath;
            public string? ArchivPath { get { return _archivPath; } set { _archivPath = value; OnPropertyChanged(nameof(ArchivPath)); } }
            private Archivator.ArchivState _archivState = archivState;
            public Archivator.ArchivState ArchivState
            {
                get { return _archivState; }
                set
                {
                    _archivState = value;
                    OnPropertyChanged(nameof(ArchivState));
                }
            }

            public event PropertyChangedEventHandler? PropertyChanged;
            private void OnPropertyChanged(string propertyName)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

    }
}
