using El2Core.Models;
using El2Core.Services;
using System;
using System.Collections.Generic;
using System.Text;

namespace ModuleProxyOrder.Entities
{
    public class OrderProxy : ProxyOrder, IDbTarget
    {
        public object? Target { get; set; }
        public State EntityState { get; private set; } = State.Unchanged;
        public OrderProxy(long OID, string AccID) { OrderId = OID; AccId = AccID; EntityState |= State.InValid; }
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
            }
            else if (!string.IsNullOrEmpty(prx.ProjId))
            {
                Target = prx.Proj;
            }
            else if (prx.CostId.HasValue)
            {
                Target = prx.Cost;
            }

        }
        public void UpdateTarget(object target)
        {
            if (target == null && Target != null)
            {
                EntityState |= State.Deleted;
            }
            else if (Target == null)
            {
                EntityState |= State.New;
            }
            else if (Target != target)
            {
                EntityState |= State.Updated;
            }
            else { EntityState |= State.Unchanged; return; }

            if (target is OrderRb rb)
            {

                RbId = rb.Aid;
                ProjId = null;
                CostId = null;
            }
            else if (target is Project pr)
            {
                ProjId = pr.ProjectPsp;
                RbId = null;
                CostId = null;
            }
            
            else if (target is Costunit cost)
            {
                CostId = cost.CostunitId;
                RbId = null;
                ProjId = null;
            }
            Target = target;
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
