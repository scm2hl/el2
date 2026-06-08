using System.Collections.Generic;

namespace El2Core.Utils
{
    public sealed class PermissionsProvider
    {
        private static PermissionsProvider? _instance;

        private HashSet<string> _permissions = new();
        private HashSet<int> _fullAccesses = new();

        public static PermissionsProvider GetInstance()
        {
            if (_instance == null)
            {
                _instance = new PermissionsProvider();
                _instance.LoadPermissions(UserInfo.User);
                return _instance;
            }
            return _instance;
        }

        private PermissionsProvider() { }

        private void LoadPermissions(User user)
        {
            foreach (var item in user.Permissions)
            {
                _permissions.Add(item);                
            }
            foreach (var access in user.AccountWorkAreas)
            {
                if(access.FullAccess)
                {
                     _fullAccesses.Add(access.WorkAreaId);
                }
            }
        }

        public bool GetUserPermission(string permission)
        {
            if (permission.StartsWith('!')) //negation of permission
            {
                return !_permissions.Contains(permission[1..]);
            }
            return _permissions.Contains(permission);
        }
        public bool GetRelativeUserPermission(string permission, int workAreaId)
        {
            if(GetUserPermission(permission))
            {
                return _fullAccesses.Contains(workAreaId);
            }
            return false;
        }
    }
}
