using Microsoft.Extensions.Logging;
using Prism.Commands;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ModuleShift.Dialogs
{
    public class ManipulationRuleEditVM : IDialogAware
    {
        private IContainerExtension _container;
        private ILogger<WeekPreviewViewModel> logger;

        public DialogCloseListener RequestClose { get; }
        public string Title { get; } = "Manipulation Rule Editor";
        public string? RuleText { get; set; }
        private DelegateCommand? _okDialogCommand;
        public DelegateCommand OkDialogCommand =>
            _okDialogCommand ?? (_okDialogCommand = new DelegateCommand(OnOkDialog));
        public ManipulationRuleEditVM(IContainerExtension container)
        {
            _container = container;
            var factory = _container.Resolve<ILoggerFactory>();
            logger = factory.CreateLogger<WeekPreviewViewModel>();
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
            RuleText = parameters.GetValue<string>("RuleText");
        }
        private void OnOkDialog()
        { 
            IDialogParameters param = new DialogParameters
            {
                { "RuleText", RuleText }
            };

            RequestClose.Invoke(param, ButtonResult.OK);
        }
    }
}
