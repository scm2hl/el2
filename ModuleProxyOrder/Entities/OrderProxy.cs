using El2Core.Models;
using El2Core.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace ModuleProxyOrder.Entities
{
    public class OrderProxy : ProxyOrder, IDbTarget, INotifyPropertyChanged, INotifyDataErrorInfo
    {
        private readonly Dictionary<string, List<string>> _errors = new();
        public event EventHandler<DataErrorsChangedEventArgs>? ErrorsChanged;
        public bool HasErrors => _errors.Any();
        public System.Collections.IEnumerable GetErrors(string? propertyName)
            => _errors.ContainsKey(propertyName) ? _errors[propertyName] : null;

        public void SetErrors(string propertyName, List<string> errors)
        {
            _errors[propertyName] = errors;
            ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
        }

        public void ClearErrors(string propertyName)
        {
            if (_errors.Remove(propertyName))
            {
                ErrorsChanged?.Invoke(this, new DataErrorsChangedEventArgs(propertyName));
            }
        }
        
        public new long OrderId
        {
            get => base.OrderId;
            set
            {
                if (base.OrderId != value)
                {
                    base.OrderId = value;
                    EntityState = State.New;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(OrderId)));
                }
            }
        }

        public object? Target
        {
            get;
            set
            {
                if (field == null)
                {
                    EntityState = State.New;
                } else if (field != value) {
                    EntityState = State.Updated;
                }
                field = value;
            }
        }
        private string _targetId;

        public string TargetId
        {
            get { return _targetId; }
            set
            {
                _targetId = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TargetId)));
            }
        }

        public State EntityState { get; private set; } = State.Unchanged;

        public event PropertyChangedEventHandler? PropertyChanged;
        
        public OrderProxy(string AccID) 
        {
            AccId = AccID;
            EntityState = State.InValid;
            Created = DateTime.Now;
            
        }
        public OrderProxy(ProxyOrder prx)
        {
            foreach (var b in prx.GetType().GetProperties())
            { 
                var value = b.GetValue(prx);
                if (value != null)
                {
                    b.SetValue(this, value);
                }
            }


            if (!string.IsNullOrEmpty(prx.RbId))
            {
                Target = prx.Rb;
                TargetId = prx.RbId;
            }
            else if (!string.IsNullOrEmpty(prx.ProjId))
            {
                Target = prx.Proj;
                TargetId = prx.Proj.ProjectPsp;
            }
            else if (prx.CostId.HasValue)
            {
                Target = prx.Cost;
                TargetId = prx.Cost.CostunitId.ToString();
            }
            EntityState = State.Unchanged;
        }
        public void UpdateTarget(object? target)
        {
            if (target == null && Target != null)
            {
                EntityState = State.Deleted;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EntityState)));
            }
            else if (Target == null)
            {
                EntityState = State.New;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EntityState)));
            }
            else if (Target != target)
            {
                EntityState = State.Updated;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EntityState)));
            }
            else { EntityState = State.Unchanged; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EntityState))); return; }

            if (target is OrderRb rb)
            {

                RbId = rb.Aid;
                ProjId = null;
                CostId = null;
                TargetId = rb.Aid;
            }
            else if (target is Project pr)
            {
                ProjId = pr.ProjectPsp;
                RbId = null;
                CostId = null;
                TargetId = pr.ProjectPsp;
            }
            
            else if (target is Costunit cost)
            {
                CostId = cost.CostunitId;
                RbId = null;
                ProjId = null;
                TargetId = cost.CostunitId.ToString();
            }
            Target = target;
        }
        public void SetState(State state)
        {
            EntityState = state;
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EntityState)));
        }   
        public enum State
        {
            Unchanged,
            New,
            Updated,
            Deleted,
            InValid
        }
    }
}
