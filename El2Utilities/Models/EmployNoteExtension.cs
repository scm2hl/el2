using El2Core.Utils;

namespace El2Core.Models
{
    public static class EmployNoteExtensions
    {
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
        public static void SetTarget(this EmployeeNote note, object? target)
        {
            if (target is EmploySelection sel)
            {
                note.SelId = sel.Id;
                note.VorgId = null;
                note.Prxid = null;
            }
            else if (target is VorgItem vorg)
            {
                note.VorgId = vorg.VorgangId;
                //note.SelId = null;
                note.SelId = null;
                note.Prxid = null;
            }
            else if (target is ProxyOrder prx)
            {
                note.Prxid = prx.OrderId;
                note.VorgId = null;
                note.SelId = null;
            }
            else
            {
                note.SelId = null;
                note.VorgId = null;
                note.Prxid = null;
            }
        }
        public static string? GetTargetString(this EmployeeNote note)
        {
            if (note.Sel != null)
                return $"{note.Sel.Description}";
            else if (note.Vorg != null)
                return $"{note.Vorg.Aid} - {note.Vorg.Vnr}\n{note.Vorg.Text}";
            else if (note.Prx != null)
                return $"{note.Prx.OrderId} - {note.Prx.CommentText}";
            return null;
        }
        public static string ToString(this EmployeeNote note)
        {
            return note.GetTargetString() ?? string.Empty;
        }
    }
}
