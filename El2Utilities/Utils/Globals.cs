using El2Core.Models;
using El2Core.Services;
using Microsoft.EntityFrameworkCore;
using Prism.Ioc;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.IO;
using System.Linq;
using static El2Core.Utils.Archivator;
using static El2Core.Utils.RuleInfo;
using Rule = El2Core.Models.Rule;

namespace El2Core.Utils

{
    public interface IGlobals
    {
        static string PC { get; }
        static User User { get; }
        static IContainerProvider Container { get; }
    }
    public class Globals : IGlobals, IDisposable
    {
        public User User { get; private set; }
        public string PC { get; }
        public List<Rule> Rules { get; private set; }
        public static NotifyBroker NotifyBroker { get; } = new NotifyBroker();
        private static IContainerProvider Container;
        public Globals(IContainerProvider container)
        {
            Container = container;
            PC = Environment.MachineName;
            
            LoadData();
            
        }
        public static UserInfo CreateUserInfo(IContainerProvider container)
        {
            //user = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string us = Environment.UserName;

            using (var db = container.Resolve<DB_COS_LIEFERLISTE_SQLContext>())
            {
                var result = (from account in db.IdmAccounts
                              join r in db.IdmRelations on account.AccountId equals r.AccountId
                              join role in db.IdmRoles on r.RoleId equals role.RoleId
                              join perm in db.RolePermissions on role.RoleId equals perm.RoleId
                              where account.AccountId == us
                              select new { account, perm }) ?? throw new KeyNotFoundException("User nicht gefunden");

                UserInfo userInfo = new UserInfo();
                var usr = result.First().account;
                var user = new User(usr.AccountId, usr.Firstname, usr.Lastname, usr.Email);
                foreach (var r in result)
                {
                    user.Permissions.Add(r.perm.PermissKey.Trim());                  
                }

                var acc = db.IdmAccounts
                    .Include(x => x.AccountCosts)
                    .ThenInclude(x => x.Cost)
                    .Include(x => x.AccountWorkAreas)
                    .ThenInclude(x => x.WorkArea)
                    .First(x => x.AccountId == us);
                user.AccountCostUnits = acc.AccountCosts.ToList();
                user.AccountWorkAreas = acc.AccountWorkAreas.ToList();
                userInfo.Initialize(Environment.MachineName, user);

                return userInfo;  
               
            }
        }
        private void LoadData()
        {

            using (var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>())
            {

                Rules = db.Rules.ToList();
            }
            
            var ext = Rules.Find(x => x.RuleName.Trim() == "Archivator");
            if (ext != null)
            {
                var serializer = XmlSerializerHelper.GetSerializer(typeof(ArchivatorWrap));
                
                TextReader xmlData = new StringReader(ext.RuleData);
                var arc = serializer.Deserialize(xmlData);
                (arc as ArchivatorWrap)?.SetArchivator();
            }
            Archivator.IsChanged = false;

            ext = Rules.Find(x => x.RuleName.Trim() == "Notification");
            if (ext != null)
            {
                var serializer = XmlSerializerHelper.GetSerializer(typeof(List<Abonnent>));
                TextReader xmlData = new StringReader(ext.RuleData);
                var abos = serializer.Deserialize(xmlData);
                if (abos == null) return;
                  
                NotifyBroker.Abonnents = (List<Abonnent>)abos;                
            }

        }
        public static void SaveRule(string RuleKey, string value)
        {
            Rule rule;
            if (RuleInfo.Rules.TryGetValue(RuleKey, out var outRule))
            {
                rule = outRule;
                rule.RuleValue = value;
            }
            else { rule = new Rule() { RuleName = RuleKey, RuleValue = value }; }
            SaveRule(rule);
        }
        public static void SaveRule(Rule rule)
        {           
            using var db = Container.Resolve<DB_COS_LIEFERLISTE_SQLContext>();
            if (db.Rules.All(x => x.RuleName != rule.RuleName))
            {
                db.Rules.Add(rule);
            }
            else
            {
                var r = db.Rules.First(x => x.RuleName == rule.RuleName);
                r.RuleValue = rule.RuleValue;
                r.RuleData = rule.RuleData;
            }
            db.SaveChanges();
        }
        public static void SaveProjectSchemes(List<ProjectScheme> projectSchemes)
        {
            var serializer = XmlSerializerHelper.GetSerializer(typeof(List<ProjectScheme>));
            StringWriter sw = new StringWriter();
            serializer.Serialize(sw, projectSchemes);

                Rule rule = new Rule();
                rule.RuleName = "ProjStruct";
                rule.RuleValue = "ProjStruct";
                rule.RuleData = sw.ToString();
            SaveRule(rule);

        }
        public static void SaveArchivator()
        {
            ArchivatorWrap wrap = new();
            wrap.GetArchivator();
            var serializer = XmlSerializerHelper.GetSerializer(typeof(ArchivatorWrap));
            StringWriter sw = new StringWriter();
            serializer.Serialize(sw, wrap);
            Rule rule = new Rule();
            rule.RuleName = "Archivator";
            rule.RuleValue = "Archivator";
            rule.RuleData = sw.ToString();
            SaveRule(rule);
            Archivator.IsChanged = false;
        }

        public static void SaveNotifyBroker()
        {
            var serializer = XmlSerializerHelper.GetSerializer(typeof(List<Abonnent>));
            StringWriter sw = new StringWriter();
            serializer.Serialize(sw, NotifyBroker.Abonnents);
            Rule rule = new Rule();
            rule.RuleName = "Notification";
            rule.RuleValue = "NotifyBroker";
            rule.RuleData = sw.ToString();
            SaveRule(rule);
        }
        public void Dispose()
        {

        }
    }
    public readonly struct UserInfo
    {
        public static string? PC => _PC ?? string.Empty;
        public static User User => _User;
        public static int Dbid;
        private static string? _PC;
        private static User _User;

        public void Initialize(string PC, User Usr)
        {
            _PC = PC;
            _User = Usr;
        }
    }
    public readonly struct RuleInfo
    {
        public static ImmutableDictionary<string, Rule> Rules { get; private set; } = ImmutableDictionary<string, Rule>.Empty;
        public static ImmutableDictionary<string, ProjectScheme> ProjectSchemes { get; private set; } = ImmutableDictionary<string, ProjectScheme>.Empty;
        public RuleInfo(List<Rule> rules)
        {
            Rules = rules.ToImmutableDictionary(x => x.RuleName.Trim(), x => x);
            Dictionary<string, ProjectScheme> sc = new Dictionary<string, ProjectScheme>();
            var scheme = rules.SingleOrDefault(x => x.RuleValue.Trim() == "ProjStruct");
            if (scheme != null)
            {
                if (scheme.RuleData == null) return;
                var serializer = XmlSerializerHelper.GetSerializer(typeof(List<ProjectScheme>));
                TextReader xmlData = new StringReader(scheme.RuleData);
                List<ProjectScheme> projectSchemes = (List<ProjectScheme>)serializer.Deserialize(xmlData);

                foreach(var item in projectSchemes)
                {
                    sc.Add(item.Key, item);
                }
                ProjectSchemes = sc.ToImmutableDictionary();
            }
        }
        public class ArchivatorWrap
        {

            public List<ArchivatorRule>? ArchiveRules { get; set; }
            public int DelayDays { get; set; }
            public string[]? FileExtensions { get; set; } 
            
            public void SetArchivator()
            {
                if (ArchiveRules != null)
                    Archivator.ArchiveRules.AddRange(ArchiveRules);

                Archivator.DelayDays = DelayDays;
                Archivator.FileExtensions = FileExtensions;
            }
            public void GetArchivator()
            {

                ArchiveRules = Archivator.ArchiveRules;
                DelayDays = Archivator.DelayDays;
                FileExtensions = Archivator.FileExtensions;
            }
        }
    }
    public class ProjectScheme
    {
        public string Key { get; set; }
        public string Regex { get; set; }
        public ProjectScheme() { }
        public ProjectScheme(string key, string regex) { Key = key; Regex = regex; }
    }
    public class User(string id, string? firstName, string? lastname, string? email)
    {
        public string? FirstName { get; } = firstName;
        public string? LastName { get; } = lastname;
        public string? Email { get; } = email;
        public string UserId { get; } = id;
        public string UsrName { get { return string.Format("{0} {1}", FirstName, LastName); } }
        public HashSet<string> Permissions { get; } = new HashSet<string>();
        public HashSet<string> Roles { get; } = [];
        public List<AccountCost> AccountCostUnits { get; set; }
        public List<AccountWorkArea> AccountWorkAreas { get; set; }

    }

}
