using MailKit.Net.Smtp;
using El2Core.Utils;
using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client.Platforms.Features.DesktopOs.Kerberos;
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

        public List<Abonnent> Abonnents { get; set; } = [];
        
        public bool GetAbonnentById(string id, out Abonnent abo)
        {
            var a = Abonnents.SingleOrDefault(m => m.Id == id);
            if (a.Name != default)
            {
                abo = a;
                return true;
            }
            else
            {
                abo = default;
                return false;
            }           
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


            try
            {
                var message = new MimeMessage();
                message.From.Add(new MailboxAddress("HlP Lieferliste", "Lieferliste.HlP@at.bosch.com"));
                
                foreach (var abo in Abonnents.Where(x => x.Subsribes.Contains(sender)))
                {
                    message.To.Add(new MailboxAddress(abo.Name, abo.Address));
                }
                if (message.To.Count > 0)
                {
                    message.Subject = "Abonnierte Nachricht";
                    var mb = message_body.Split((char)29);
                    message.Body = new TextPart("html") { Text = $"<b>Hallo! Abonnents<p>Nachricht von {sender.Description()}</p></b><p>{mb[0]} hat folgendes geschrieben.</p>{mb[1]}" };

                    using (var client = new MailKit.Net.Smtp.SmtpClient())
                    {

                        await client.ConnectAsync("smtp.app.bosch.com", 587, MailKit.Security.SecureSocketOptions.Auto);

                        client.Authenticate("HLS2HL@bosch.com", "d0K1nzTUj85x");
                        await client.SendAsync(message);
                        await client.DisconnectAsync(true);
                    }

                }
            
            }
            catch (System.Exception e)
            {

                throw;
            }
        }

        public bool UpdateAbonnent(Abonnent value)
        {
            var abo = Abonnents.SingleOrDefault(m => m.Id == value.Id);
            if (abo.Name != default)
            {
                if (abo.Subsribes == value.Subsribes)
                {
                    return false;
                }
                else
                {
                    abo.Address = value.Address;
                    abo.Subsribes = value.Subsribes;
                    return true;
                }
            }
            else  return false;
        }
    }
}
