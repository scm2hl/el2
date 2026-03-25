using El2Core.Utils;
using Microsoft.Extensions.Configuration;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

using System.Threading.Tasks;

namespace El2Core.Services
{
    public interface INotifyBroker
    {
        List<Abonnent> Abonnents { get; set; }
        bool GetAbonnentById(string id, out Abonnent abo);
        Task SendMessageAsync(string message_body, SubscribeType sender);
        void AddAbonnent(Abonnent abonnent);
        bool RemoveAbonnent(Abonnent abonnent);
        bool UpdateAbonnent(Abonnent value);
    }
    public struct Abonnent
    {
        public string Id { get; set; }
        public string Address { get; set; }
        public string Name { get; set; }
        public SubscribeType[] Subsribes { get; set; }
        
    }
    public enum SubscribeType
    {
        [Description("ohne")]
        None = 0,
        [Description("Bemerkung Meister")]
        MeBem = 1,
        [Description("Bemerkung Teamleiter")]
        TeBem = 2,
        [Description("Bemerkung Mitarbeiter")]
        MaBem = 3
    }
    public class NotifyBroker : INotifyBroker
    {
        private static IConfiguration? _container;
        public List<Abonnent> Abonnents { get; set; } = new List<Abonnent>();
        
        public bool GetAbonnentById(string id, out Abonnent abo)
        {
            var idx = Abonnents.FindIndex(a => a.Id == id);
            if (idx >= 0)
            {
                abo = Abonnents[idx];
                return true;
            }

            abo = default;
            return false;
        }

        public void AddAbonnent(Abonnent abonnent)
        {
            Abonnents.Add(abonnent);
        }

        public bool RemoveAbonnent(Abonnent abonnent)
        {
            return Abonnents.Remove(abonnent);
        }

        public async Task SendMessageAsync(string message_body, SubscribeType sender)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("HlP Lieferliste", "Lieferliste.HlP@at.bosch.com"));

            // Add recipients who subscribed to this type
            foreach (var abo in Abonnents.Where(x => x.Subsribes != null && x.Subsribes.Contains(sender)))
            {
                message.To.Add(new MailboxAddress(abo.Name, abo.Address));
            }

            if (message.To.Count == 0)
                return;

            message.Subject = "Abonnierte Nachricht";
            var mb = message_body?.Split((char)29) ?? Array.Empty<string>();
            var first = mb.Length > 0 ? mb[0] : string.Empty;
            var second = mb.Length > 1 ? mb[1] : string.Empty;
            var third = mb.Length > 2 ? mb[2] : string.Empty;
            message.Body = new TextPart("html") { Text = $"<b>Hallo! Abonnents<p>Nachricht von {sender.Description()}</p></b><p>{first} hat folgendes geschrieben.</p>" +
                $"<p>{second}</p><p></p>Referenz: {third}" };

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                await client.ConnectAsync("smtp.app.bosch.com", 587, MailKit.Security.SecureSocketOptions.Auto);
                client.Authenticate("HLS2HL@bosch.com", "d0K1nzTUj85x");
                await client.SendAsync(message);
                await client.DisconnectAsync(true);
            }
        }

        public bool UpdateAbonnent(Abonnent value)
        {
            var idx = Abonnents.FindIndex(m => m.Id == value.Id);
            if (idx < 0)
                return false;

            var existing = Abonnents[idx];
            // If subscriptions are the same (reference or sequence), nothing to do
            if (existing.Subsribes == value.Subsribes || (existing.Subsribes != null && value.Subsribes != null && existing.Subsribes.SequenceEqual(value.Subsribes)))
                return false;

            // Replace the item in the list (Abonnent is a struct)
            Abonnents[idx] = value;
            return true;
        }
    }
}
