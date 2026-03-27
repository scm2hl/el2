using CompositeCommands.Core;
using El2Core.Utils;
using GongSolutions.Wpf.DragDrop;
using ModulePlanning.Planning;
using Prism.Dialogs;
using System.Threading;
using System.Windows.Input;

namespace ModulePlanning.Dialogs.ViewModels
{
    [System.Runtime.Versioning.SupportedOSPlatform("windows10.0")]
    public class MachineViewVM : IDialogAware, IDropTarget
    {
        public string Title => "Maschinen Details";
        public string UserIdent { get; } = UserInfo.User.UserId;
        public PlanMachine? PlanMachine { get; private set; }
        private IApplicationCommands _applicationCommands;
        public IApplicationCommands ApplicationCommands
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
        public SynchronizationContext ViewContext { get; set; }
        public ICommand? SetMarkerCommand { get; private set; }
        public ICommand? HistoryCommand { get; private set; }
        public ICommand? FastCopyCommand { get; private set; }
        public ICommand? CorrectionCommand { get; private set; }

        public DialogCloseListener RequestClose { get; }

        public MachineViewVM(IApplicationCommands applicationCommands) { _applicationCommands = applicationCommands; }
        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {
           
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            PlanMachine = parameters.GetValue<PlanMachine>("PlanMachine");
            SetMarkerCommand = PlanMachine.SetMarkerCommand;
            HistoryCommand = PlanMachine.HistoryCommand;  
            FastCopyCommand = PlanMachine.FastCopyCommand;
            CorrectionCommand = PlanMachine.CorrectionCommand;
        }
        public void Drop(IDropInfo dropInfo)
        {
            PlanMachine?.Drop(dropInfo);
        }

        public void DragOver(IDropInfo dropInfo)
        {
            if (dropInfo.IsSameDragDropContextAsSource) PlanMachine?.DragOver(dropInfo);

        }
    }
}
