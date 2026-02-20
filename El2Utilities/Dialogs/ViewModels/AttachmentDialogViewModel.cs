using El2Core.Constants;
using El2Core.Models;
using El2Core.Utils;
using GongSolutions.Wpf.DragDrop;
using Microsoft.Extensions.Logging;
using Microsoft.Win32;
using Prism.Dialogs;
using Prism.Ioc;
using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;


namespace El2Core.Dialogs.ViewModels
{

    public class AttachmentDialogViewModel : IDropTarget, IDialogAware
    {

        public AttachmentDialogViewModel(IContainerProvider container)
        {
            Container = container;
            var loggerFactory = container.Resolve<ILoggerFactory>();
            Logger = loggerFactory.CreateLogger<AttachmentDialogViewModel>();
            OpenFileCommand = new ActionCommand(OnFileOpenExecuted, OnOpenFileCanExecute);
            DelAttachmentCommand = new ActionCommand(OnDelAttachmentExecuted, OnDelAttachmentCanExecute);
            AddLinkedAttachmentCommand = new ActionCommand(OnLinkedAttachmentExecuted, OnLinkedAttachmentCanExecute);
            AddAttachmentCommand = new ActionCommand(OnAttachmentExecuted, OnAttachmentCanExecute);
        }

        public string Title { get; } = "Anhang Vorgang";
        private Vorgang Vorgang { get; set; }

        public DialogCloseListener RequestClose { get; }

        IContainerProvider Container;
        ILogger Logger;
        private RelayCommand? _closeCommand;
        public RelayCommand CloseCommand => _closeCommand ??= new RelayCommand(OnDialogClosed);

        private void OnDialogClosed(object obj)
        {
            OnDialogClosed();
        }

        public ICommand OpenFileCommand { get; private set; }
        public ICommand DelAttachmentCommand { get; private set; }
        public ICommand AddLinkedAttachmentCommand { get; private set; }
        public ICommand AddAttachmentCommand { get; private set; }
        public ICollectionView AttachView { get; private set; }
        private ObservableCollection<IDisplayAttachment> _attachments = [];
        private bool OnDelAttachmentCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.DelVrgAttachment);
        }

        private void OnDelAttachmentExecuted(object obj)
        {
            if (obj is VrgDisplayAttachment disp)
            {
                var att = _attachments.Single(x => x.Id == disp.Id);
                _attachments.Remove(att);

                using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                var dbatt = db.VorgangAttachments.Single(x => x.AttachId == disp.Id);
                db.VorgangAttachments.Remove(dbatt);
                db.SaveChanges();
                Logger.LogInformation("{message}", dbatt.Link);
            }
        }
        private bool OnLinkedAttachmentCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVrgAttachment);
        }

        private async void OnLinkedAttachmentExecuted(object obj)
        {

            var f = await AttachmentFactory.GetFilePickerPath();
            if (string.IsNullOrEmpty(f))
            {
                AddAttachment(f, true);
                Logger.LogInformation("{message}", f);
            }
        }
        private bool OnAttachmentCanExecute(object arg)
        {
            return PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVrgAttachment);
        }

        private async void OnAttachmentExecuted(object obj)
        {

            var f = await AttachmentFactory.GetFilePickerPath();
            if (!string.IsNullOrEmpty(f))
            {
                AddAttachment(f, false);
                Logger.LogInformation("{message}", f);
            }
        }
        private bool OnOpenFileCanExecute(object arg)
        {
            return true;
        }

        private void OnFileOpenExecuted(object obj)
        {
            if (obj is VrgDisplayAttachment disp)
            {
                var fact = new VorgangAttachmentCreator();
                if (disp.IsLink)
                {
                    AttachmentFactory.OpenFile(disp.Name, null);
                }
                else
                {
                    using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
                    var att = db.VorgangAttachments.Single(x => x.AttachId.Equals(disp.Id));
                    using MemoryStream ms = new(att.Data);
                    AttachmentFactory.OpenFile(disp.Name, ms);
                }
            }
        }
        public void DragOver(IDropInfo dropInfo)
        {
            if (PermissionsProvider.GetInstance().GetUserPermission(Permissions.AddVrgAttachment))
            {
                dropInfo.DropTargetAdorner = DropTargetAdorners.Highlight;
                dropInfo.Effects = DragDropEffects.All;
            }
        }

        public void Drop(IDropInfo dropInfo)
        {
            if (dropInfo.Data is IDataObject f)
            {
                var o = (string[])f.GetData(DataFormats.FileDrop);
                if (o.Length > 0)
                {
                    AddAttachment(o[0], false);
                    Logger.LogInformation("Att droped {message}", o[0]);
                }
            }
        }

        public void OnDialogClosed()
        {
            IDialogParameters param = new DialogParameters();
            RequestClose.Invoke(param);
        }

        public void OnDialogOpened(IDialogParameters parameters)
        {
            using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var vrg = parameters.GetValue<Vorgang>("vrg");
            this.Vorgang = vrg;
            var vaFactory = new VorgangAttachmentCreator();

            var attach = db.VorgangAttachments.Where(x => x.VorgangId == vrg.VorgangId).ToList();
            foreach (var att in attach)
            {
                var va = vaFactory.CreateDisplayAttachment(att.Link, att.IsLink);
                va.Id = att.AttachId;
                _attachments.Add(va);
            }
            AttachView = CollectionViewSource.GetDefaultView(_attachments);
        }

        public bool CanCloseDialog()
        {
            return true;
        }
        private void AddAttachment(string link, bool isLink)
        {
            var dbfact = new VorgangAttachmentCreator();

            var att = dbfact.CreateDbAttachment(link, isLink);

            var vatt = new VorgangAttachment();
            vatt.Timestamp = att.TimeStamp;
            vatt.Data = att.BinaryData;
            vatt.Link = att.Link;
            vatt.IsLink = att.IsLink;
            using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            var vrg = db.Vorgangs.Single(x => x.VorgangId == this.Vorgang.VorgangId);
            vrg.VorgangAttachments.Add(vatt);
            db.SaveChanges();


            var attDisp = dbfact.CreateDisplayAttachment(att.Link, att.IsLink);
            attDisp.Id = vatt.AttachId;
            _attachments.Add(attDisp);
        }


    }
    internal class VrgDisplayAttachment : IDisplayAttachment
    {

        public string Name { get; set; } = string.Empty;
        public string? Description { get; set; }
        public object? Content { get; set; }
        public int Id { get; set; }
        public bool IsLink { get; set; }
    }
    internal class VrgDbAttachment : IDbAttachment
    {
        public string Link { get; set; } = string.Empty;
        public bool IsLink { get; set; }
        public DateTime TimeStamp { get; set; }
        public byte[]? BinaryData { get; set; }
    }
    internal class VorgangAttachmentCreator : AttachmentFactory
    {
        public override IDbAttachment CreateDbAttachment(string Link, bool isLink)
        {
            var attachment = new VrgDbAttachment();
            var att = FloatAttachment(attachment, Link, isLink);
            return att;
        }

        public override IDisplayAttachment CreateDisplayAttachment(string link, bool isLink)
        {
            
            var attachment = new VrgDisplayAttachment();
            var att = FloatAttachment(attachment, link, isLink);
            return att;
        }
    }
}
