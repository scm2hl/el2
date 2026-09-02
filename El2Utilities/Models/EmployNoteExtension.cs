using El2Core.Utils;

namespace El2Core.Models
{
    public static class EmployNoteExtensions
    {
        /// <summary>
        /// Gets the target object (EmploySelection, VorgItem, or ProxyOrder) associated with the EmployeeNote.
        /// </summary>
        /// <param name="note"></param>
        /// <returns></returns>
        public static object? GetTarget(this EmployeeNote note)
        {
            object? ret = null;
            if (note.Sel != null)
                ret = note.Sel;
            else if (note.Vorg != null)
                ret = note.Vorg;
            else if (note.Prx != null)
                ret = note.Prx;
            return ret;
        }
        /// <summary>
        /// Sets the target object (EmploySelection, VorgItem, or ProxyOrder) for the EmployeeNote.
        /// </summary>
        /// <param name="note"></param>
        /// <param name="target"></param>
        public static void SetTarget(this EmployeeNote note, object? target)
        {
            if (target is EmploySelection sel)
            {
                note.SelId = sel.Id;
                note.VorgId = null;
                note.PrxId = null;
            }
            else if (target is VorgItem vorg)
            {
                note.VorgId = vorg.VorgangId;
                //note.SelId = null;
                note.SelId = null;
                note.PrxId = null;
            }
            else if (target is ProxyOrder prx)
            {
                note.PrxId = prx.OrderId;
                note.VorgId = null;
                note.SelId = null;
            }
            else
            {
                note.SelId = null;
                note.VorgId = null;
                note.PrxId = null;
            }
        }
        /// <summary>
        /// Gets a string representation of the target object associated with the EmployeeNote.
        /// </summary>
        /// <param name="note"></param>
        /// <returns></returns>
        public static string? GetTargetInfo(this EmployeeNote note)
        {
            if (note.Sel != null)
                return $"{note.Sel.Description}";
            else if (note.Vorg != null)
                return $"{note.Vorg.Aid} - {note.Vorg.Vnr}\n{note.Vorg.Text}\n{note.Vorg.AidNavigation?.Material} - {note.Vorg.AidNavigation?.MaterialNavigation?.Bezeichng}";
            else if (note.Prx != null)
                return $"{note.Prx.OrderId} - {note.Prx.CommentText}";
            return null;
        }
 
    }

}
