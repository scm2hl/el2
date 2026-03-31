using El2Core.Models;
using El2Core.ViewModelBase;
using Microsoft.EntityFrameworkCore.Diagnostics;
using Microsoft.Extensions.Logging;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Data;

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
        private ObservableCollection<ProxyOrder> _proxies;
        public ICollectionView Proxies { get; private set; }

        private void LoadData()
        {
            using var db = _containerExtension.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var prx = db.ProxyOrders.OrderByDescending(x => x.Created).ToList();
            Proxies = CollectionViewSource.GetDefaultView(prx);
        }
    }
}
