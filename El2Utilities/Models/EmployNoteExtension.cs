using El2Core.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace El2Core.Models
{
    public static class EmployNoteExtensions
    {
        public static object GetTarget(this EmployeeNote note)
        {
            return note.SelId != null ? (object)note.Sel : note.VorgId != null ? (object)note.Vorg : null;
        }
        public static void SetTarget(this EmployeeNote note, object? target)
        {
            if (target is EmploySelection sel)
            {
                note.SelId = sel.Id;
                note.VorgId = null;
            }
            else if (target is VorgItem vorg)
            {
                note.VorgId = vorg.VorgangId;
                note.SelId = null;
            }
            else if (target is ProxyOrder prx)
            {
                //note.VorgId = prx.OrderId;
                note.SelId = null;
            }
            else 
            {
                note.SelId = null;
                note.VorgId = null;
            }
        }
    }
}
