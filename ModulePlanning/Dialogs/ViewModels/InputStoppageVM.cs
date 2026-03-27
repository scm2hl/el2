using El2Core.Models;
using El2Core.Utils;
using Prism.Dialogs;
using System;

namespace ModulePlanning.Dialogs.ViewModels
{
    public class InputStoppageVM : IDialogAware
    {
        public DateTime? StartDateTime { get; set; }
        public DateTime? EndDateTime { get; set; }
        public DialogCloseListener RequestClose { get; }
        public required Stopage Stopage { get; set; }
        private RelayCommand? _closeCommand;
        public RelayCommand CloseCommand => _closeCommand ??= new RelayCommand(OnDialogClosing);

        private void OnDialogClosing(object obj)
        {
            if ((string)obj == "1")
            {
                OnDialogClosed();
            }
            else RequestClose.Invoke(ButtonResult.Cancel);
        }

        public bool CanCloseDialog()
        {
            return true;
        }

        public void OnDialogClosed()
        {
            if (StartDateTime != null) Stopage.Starttime = StartDateTime.Value;
            if (EndDateTime != null) Stopage.Endtime = EndDateTime.Value;
            ButtonResult result = (StartDateTime != null && EndDateTime != null && (string.IsNullOrWhiteSpace(Stopage.Description) == false))
                ? ButtonResult.OK : ButtonResult.None;
            DialogParameters par = [];
            par.Add("Stopage", Stopage);
            RequestClose.Invoke(par, result);
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            
            var stop = parameters.GetValue<Stopage>("Stopage");
            if (stop != null)
            {
                Stopage = stop;
                StartDateTime = Stopage.Starttime;
                EndDateTime = Stopage.Endtime;
            }
            else { Stopage = new() { Description = string.Empty }; }

        }

    }
}
