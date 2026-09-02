using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Threading;

namespace El2Core.Utils
{
    /// <summary>
    /// An ObservableCollection that is thread-safe and can be updated from any thread.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ConcurrentObservableCollection<T> : ObservableCollection<T>
    {
        private readonly SynchronizationContext _synchronizationContext = SynchronizationContext.Current;

        private bool _suppressNotification;

        public ConcurrentObservableCollection()
        {
        }
        public ConcurrentObservableCollection(IEnumerable<T> list)
            : base(list)
        {
        }

        public void AddRange(IEnumerable<T> collection)
        {
            if (collection == null) return;
            _suppressNotification = true;
            foreach (var item in collection)
            {
                this.Add(item);
            }
            _suppressNotification = false;

            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));

        }
        public void RemoveRange(IEnumerable<T> collection)
        {

            _suppressNotification = true;
            foreach (var item in collection)
            {
                this.Remove(item);
            }
            _suppressNotification = false;

            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));

        }

        protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
        {
            if (SynchronizationContext.Current == _synchronizationContext)
            {
                // Execute the CollectionChanged event on the current thread
                RaiseCollectionChanged(e);
            }
            else
            {
                // Raises the CollectionChanged event on the creator thread
                _synchronizationContext.Send(RaiseCollectionChanged, e);
            }
        }
        protected override void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            if (SynchronizationContext.Current == _synchronizationContext)
            {
                // Execute the PropertyChanged event on the current thread
                RaisePropertyChanged(e);
            }
            else
            {
                // Raises the PropertyChanged event on the creator thread
                _synchronizationContext.Send(RaisePropertyChanged, e);
            }
        }

        private void RaiseCollectionChanged(object param)
        {
            // We are in the creator thread, call the base implementation directly
            if (!_suppressNotification)
                base.OnCollectionChanged((NotifyCollectionChangedEventArgs)param);
        }
        private void RaisePropertyChanged(object param)
        {
            // We are in the creator thread, call the base implementation directly
            base.OnPropertyChanged((PropertyChangedEventArgs)param);
        }
    }
}
