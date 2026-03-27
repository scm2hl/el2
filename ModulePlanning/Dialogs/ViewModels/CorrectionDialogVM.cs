using El2Core.Models;
using Prism.Commands;
using Prism.Dialogs;

namespace ModulePlanning.Dialogs.ViewModels
{
    public class CorrectionDialogVM : IDialogAware
    {
        private double? correctValue;
        public double? CorrectValue
        {
            get { return correctValue; }
            set { correctValue = value; }
        }
        private Vorgang? vorgang;
        public string Title => "Zeit Korrektur";

        private DelegateCommand<string?>? _closeDialogCommand;
        public DelegateCommand<string?> CloseDialogCommand =>
            _closeDialogCommand ?? (_closeDialogCommand = new DelegateCommand<string?>(CloseDialog));

        public DialogCloseListener RequestClose { get; }

        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {

        }
        protected virtual void CloseDialog(string? parameter)
        {
            ButtonResult result = ButtonResult.None;
            IDialogParameters param = new DialogParameters();

            if (parameter == null)
                result = ButtonResult.Cancel;
            else
            {
                result = ButtonResult.OK;
                param.Add("correct", correctValue * 60);
                param.Add("correction", vorgang);
            }
            RequestClose.Invoke(param, result);
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            vorgang = parameters.GetValue<Vorgang>("correction");
            correctValue = vorgang.Correction / 60;
            
        }
    }

}
