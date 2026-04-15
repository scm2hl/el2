using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ModuleProxyOrder.Entities;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using static ModuleProxyOrder.Entities.OrderProxy;


namespace ModuleProxyOrder.ViewModels
{
    public class ProxyOrderViewModel : ViewModelBase
    {
        public ProxyOrderViewModel(IContainerExtension container)
        {
            _containerExtension = container;
            var loggerFactory = container.Resolve<ILoggerFactory>();
            _Logger = loggerFactory.CreateLogger<ProxyOrderViewModel>();
            LoadData();
        }
        public string Title { get; } = "Ungeplante Aufträge";
        private readonly IContainerExtension _containerExtension;
        private readonly ILogger _Logger;
        private ObservableCollection<OrderProxy> _proxies = [];


        public ICollectionView Proxies { get; private set; }
        public ICollectionView Projects { get; private set; }
        public ICollectionView CostUnits { get; private set; }
        public ICollectionView Orders { get; private set; }
        public ICommand NewProxyCommand => new ActionCommand(OnNewExecuted, OnNewCanExecute);
        public ICommand DeleteProxyCommand => new ActionCommand(OnDeleteExecuted, OnDeleteCanExecute);

        public ICommand ApplyProxyCommand => new ActionCommand(OnApplyExecuted, OnApplyCanExecute);

  
        public ICommand CancelEditCommand => new ActionCommand(OnCancelExecuted);

        private void OnCancelExecuted(object obj)
        {
            Proxies.MoveCurrentToFirst();
            Current = Proxies.CurrentItem as OrderProxy;
            IsAddingNew = false;
        }

        private bool OnApplyCanExecute(object arg)
        {
            return Current != null && IsAddingNew && Current.EntityState != State.InValid
                && Current.Target != null && Current.OrderId > 0 && !Current.HasErrors;
        }

        private void OnApplyExecuted(object obj)
        {
            _proxies.Add(Current);
            IsAddingNew = false;
            Proxies.Refresh();
            SaveChanges();
        }
        private bool OnDeleteCanExecute(object arg)
        {
            return Current != null && !IsAddingNew;
        }

        private void OnDeleteExecuted(object obj)
        {
            Current.SetState(State.Deleted);
            SaveChanges();
            _proxies.Remove(Current);
            Proxies.Refresh();
        }
        private bool OnNewCanExecute(object arg)
        {
            return !IsAddingNew;
        }

        private void OnNewExecuted(object obj)
        {
            IsAddingNew = true;
            Current = new OrderProxy(UserInfo.User.UserId);
        }
        private OrderProxy _current;
        public OrderProxy Current
        {
            get { return _current; }
            set
            {
                if (_current != null)
                {
                    _current.PropertyChanged -= CurrentOnPropertyChanged;
                }

                _current = value;

                if (_current != null)
                {
                    _current.PropertyChanged += CurrentOnPropertyChanged;
                }

                NotifyPropertyChanged(() => Current);
            }
        }

        private void CurrentOnPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            // OrderProxy raises PropertyChanged for "OrderId"
            if (e.PropertyName == nameof(OrderProxy.OrderId) || e.PropertyName == "OrderId")
            {
                ValidateOrderId();
            }
        }
        private object? _selectedTarget;
        public object? SelectedTarget
        {  get { return _selectedTarget; } 
            set
            {
                if (value == null) { _selectedTarget = null; return; }
                _selectedTarget = value;

                    Current.UpdateTarget(SelectedTarget);
                    NotifyPropertyChanged(() => SelectedTarget);
                

            }
        }
        private string _SearchText;

        public string SearchText
        {
            get { return _SearchText; }
            set
            {
                _SearchText = value;
                Proxies.Refresh();
            }
        }
        private string _ProjectSearchText;

        public string ProjectSearchText
        {
            get { return _ProjectSearchText; }
            set
            {
                _ProjectSearchText = value;
                Projects.Refresh();
            }
        }
        private string _CostSearchText;

        public string CostSearchText
        {
            get { return _CostSearchText; }
            set
            {
                _CostSearchText = value;
                CostUnits.Refresh();
            }
        }
        private string _OrderSeachText;

        public string OrderSearchText
        {
            get { return _OrderSeachText; }
            set
            {
                _OrderSeachText = value;
                Orders.Refresh();
            }
        }


        private bool _isAddingNew;
        public bool IsAddingNew
        {
            get => _isAddingNew;
            set
            {
                _isAddingNew = value;
                NotifyPropertyChanged(() => IsAddingNew);
            }
        }

        private void LoadData()
        {
            using var db = _containerExtension.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var prx = db.ProxyOrders.OrderByDescending(x => x.Created).Include(x => x.Rb).Include(x => x.Proj).Include(x => x.Cost).ToList();

            foreach (var item in prx)
            {
                _proxies.Add(new OrderProxy(item));
            }

            Proxies = CollectionViewSource.GetDefaultView(_proxies);
            Projects = CollectionViewSource.GetDefaultView(db.Projects.ToList());
            CostUnits = CollectionViewSource.GetDefaultView(db.Costunits.ToList());
            Orders = CollectionViewSource.GetDefaultView(db.OrderRbs.Where(x => x.Abgeschlossen == false).ToList());
            Projects.Filter = (x) => string.IsNullOrEmpty(ProjectSearchText) || (x as Project)?.ProjectPsp.Contains(ProjectSearchText, StringComparison.InvariantCultureIgnoreCase) == true;
            CostUnits.Filter = (x) => string.IsNullOrEmpty(CostSearchText) || ((x as Costunit)?.CostunitId.ToString().Contains(CostSearchText, StringComparison.InvariantCultureIgnoreCase) == true);
            Orders.Filter = (x) => string.IsNullOrEmpty(OrderSearchText) || ((x as OrderRb)?.Aid.Contains(OrderSearchText, StringComparison.InvariantCultureIgnoreCase) == true);
            Proxies.Filter = (x) => string.IsNullOrEmpty(SearchText) ||
                ((x as OrderProxy)?.OrderId.ToString().Contains(SearchText, StringComparison.InvariantCultureIgnoreCase) == true) ||
                ((x as OrderProxy)?.CommentText?.Contains(SearchText, StringComparison.InvariantCultureIgnoreCase) == true);

            if (Proxies.IsEmpty)
            {
                IsAddingNew = true;
                Current = new OrderProxy(UserInfo.User.UserId);
                Current.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(Current.OrderId))
                    {
                        ValidateOrderId();
                    }
                };

            }
            else
            {
                Current = Proxies.CurrentItem as OrderProxy;
            }

        }
        public void SaveChanges()
        {
            using var db = _containerExtension.Resolve<DB_COS_LIEFERLISTE_SQLContext>();

            
            foreach (var item in _proxies.ToList())
            {
                if (item.EntityState == State.New)
                {
                    var prx = new ProxyOrder
                    {
                        // copy scalar properties
                        OrderId = item.OrderId,
                        RbId = item.RbId,
                        CostId = item.CostId,
                        AccId = item.AccId,
                        Quantity = item.Quantity,
                        CommentText = item.CommentText,
                        Created = item.Created,
                        ProjId = item.ProjId
                    };

                    db.ProxyOrders.Add(prx);
                }
                else if (item.EntityState == State.Updated)
                {
                    var existing = db.ProxyOrders.Find(item.OrderId);
                    if (existing != null)
                    {
                        existing.RbId = item.RbId;
                        existing.CostId = item.CostId;
                        existing.AccId = item.AccId;
                        existing.Quantity = item.Quantity;
                        existing.CommentText = item.CommentText;
                        existing.Created = item.Created;
                        existing.ProjId = item.ProjId;
                        db.ProxyOrders.Update(existing);
                    }
                    else
                    {
                        // If not found in DB, add as new
                        db.ProxyOrders.Add(new ProxyOrder
                        {
                            OrderId = item.OrderId,
                            RbId = item.RbId,
                            CostId = item.CostId,
                            AccId = item.AccId,
                            Quantity = item.Quantity,
                            CommentText = item.CommentText,
                            Created = item.Created,
                            ProjId = item.ProjId
                        });
                    }
                }
                else if (item.EntityState == State.Deleted)
                {
                    var existing = db.ProxyOrders.Find(item.OrderId);
                    if (existing != null)
                    {
                        db.ProxyOrders.Remove(existing);
                    }
                    else
                    {
                        // If it was never persisted, just remove from collection
                        _proxies.Remove(item);
                        continue;
                    }
                }

                item.SetState(State.Unchanged);
            }

            db.SaveChanges();
        }

        private void ValidateOrderId()
        {
            if (Current == null)
                return;

            // Ensure any previous errors are cleared on the item
            Current.ClearErrors(nameof(Current.OrderId));

            if (_proxies.Any(p => p.OrderId == Current.OrderId))
            {
                Current.SetErrors(nameof(Current.OrderId), new List<string> { "Nummer ist schon vorhanden." });
            }
        }
    }
  
}
