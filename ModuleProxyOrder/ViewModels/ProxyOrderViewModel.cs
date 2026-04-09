using El2Core.Models;
using El2Core.Utils;
using El2Core.ViewModelBase;
using Microsoft.Extensions.Logging;
using ModuleProxyOrder.Entities;
using Prism.Ioc;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using static ModuleProxyOrder.Entities.OrderProxy;


namespace ModuleProxyOrder.ViewModels
{
    internal class ProxyOrderViewModel : ViewModelBase
    {
        public ProxyOrderViewModel(IContainerExtension container)
        {
            _containerExtension = container;
            var loggerFactory = container.Resolve<ILoggerFactory>();
            _Logger = loggerFactory.CreateLogger<ProxyOrderViewModel>();
            LoadData();
        }

        private readonly IContainerExtension _containerExtension;
        private readonly ILogger _Logger;
        private ObservableCollection<OrderProxy> _proxies = [];
        
        public ICollectionView Proxies { get; private set; }
        public ICollectionView Projects { get; private set; }
        public ICollectionView CostUnits { get; private set; }
        private object? _selectedTarget;
        public object? SelectedTarget
        {  get { return _selectedTarget; } 
            set
            {
                if (value == null) { _selectedTarget = null; return; }
                _selectedTarget = value;
                if (Proxies.CurrentItem is OrderProxy prx)
                {
                    prx.UpdateTarget(value);
                    Proxies.Refresh();
                }
            }
        }

        private readonly long _newOid;

        public long NewOid
        {
            get { return _newOid; }
            set 
            {
                _proxies.Add(new OrderProxy(value, UserInfo.User.UserId));
                IsAddingNew = true;
                    Proxies.MoveCurrentToLast();
                    Proxies.Refresh();
            }
        }
        
        public bool IsAddingNew { get; private set; }
        private void LoadData()
        {
            using var db = _containerExtension.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var prx = db.ProxyOrders.OrderByDescending(x => x.Created);
            
            foreach (var item in prx)
            {
                _proxies.Add(new OrderProxy(item));
            }
 
            Proxies = CollectionViewSource.GetDefaultView(prx);
            Projects = CollectionViewSource.GetDefaultView(db.Projects);
            CostUnits = CollectionViewSource.GetDefaultView(db.Costunits);
        }
        public void SaveChanges()
        {
            using var db = _containerExtension.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            foreach (var item in _proxies)
            {
                if (item.EntityState.HasFlag(State.New))
                {
                    db.ProxyOrders.Add(item);
                }
                else if (item.EntityState.HasFlag(State.Updated))
                {
                    db.ProxyOrders.Update(item);
                }
                else if (item.EntityState.HasFlag(State.Deleted))
                {
                    db.ProxyOrders.Remove(item);
                }
            }
            db.SaveChanges();
        }
    }
}
