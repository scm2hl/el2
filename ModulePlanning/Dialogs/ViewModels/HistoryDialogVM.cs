using CompositeCommands.Core;
using El2Core.Models;
using Prism.Commands;
using Prism.Dialogs;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;

namespace ModulePlanning.Dialogs.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public class HistoryDialogVM : IDialogAware
    {
        public HistoryDialogVM(IApplicationCommands applicationCommands) { _applicationCommands = applicationCommands; }
        public string Title => "Historische Aufträge";
        public string? Material => _material;

        private string? _vid;
        private string? _material;
        public string? MatDescription => _matDescription;
        private string? _matDescription;
        public int VrgNr => _vrgNr;
        private int _vrgNr;
        public IApplicationCommands? ApplicationCommands
        {
            get { return _applicationCommands; }
            set
            {
                if (_applicationCommands != value)
                {
                    _applicationCommands = value;
                }
            }
        }
        private IApplicationCommands? _applicationCommands;
        private List<Vorgang>? _orderList;
        public ICollectionView? OrderList { get; private set; }
        
        //public event Action<IDialogResult>? RequestClose;
        private DelegateCommand<Vorgang?>? _closeDialogCommand;
        public DelegateCommand<Vorgang?> CloseDialogCommand =>
            _closeDialogCommand ?? (_closeDialogCommand = new DelegateCommand<Vorgang?>(CloseDialog));

        public DialogCloseListener RequestClose { get; }

        protected virtual void CloseDialog(Vorgang? parameter)
        {
            ButtonResult result = ButtonResult.None;
            IDialogParameters param = new DialogParameters();

            if (parameter == null)
                result = ButtonResult.Cancel;
            else if (parameter is Vorgang v)
            {
                result = ButtonResult.Yes;
                param.Add("Comment", v.BemT);
                param.Add("VID", _vid);
            }
            RequestClose.Invoke(param, result);
        }

        public bool CanCloseDialog()
        {
            return true;
        }
        public void OnDialogClosed()
        {
 
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            var ord = parameters.GetValue<List<OrderRb>>("orderList");
            var vnr = parameters.GetValue<short>("VNR");
            _vid = parameters.GetValue<string>("VID");

            _material = ord.First().Material;
            _matDescription = ord.First().MaterialNavigation?.Bezeichng;
            _vrgNr = vnr;
            var list = new List<Vorgang>();
            foreach (var o in ord.OrderByDescending(x => x.Eckende))
            {
                var oo = o.Vorgangs.OrderBy(x => x.Vnr);
                for(int i = 0; i < oo.Count(); i++)
                {
                    if (oo.ElementAt(i).Vnr >= vnr)
                    {
                        
                        var e = oo.ElementAtOrDefault(i - 1);
                        if (e != null) list.Add(e);
                        list.Add(oo.ElementAt(i));
                        e = oo.ElementAtOrDefault(i  + 1);
                        if (e != null) list.Add(e);
                        break;
                    }
                }
            }
            _orderList = new List<Vorgang>(list);
            OrderList = CollectionViewSource.GetDefaultView(_orderList);
            OrderList.GroupDescriptions.Add(new PropertyGroupDescription("Aid"));
        }
    }
}
