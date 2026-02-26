using MailKit.Net.Smtp;
using Microsoft.Extensions.Configuration;
using MimeKit;
using Prism.Ioc;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace El2Core.Services
{
    public interface INotifyBroker
    {
        List<Abonnent> Abonnents { get; set; }
        bool GetAbonnentById(string id, out Abonnent abo);
        Task SendMessageAsync(string message_body, string sender);
        void AddAbonnent(Abonnent abonnent);
        bool RemoveAbonnent(Abonnent abonnent);
        bool UpdateAbonnent(Abonnent value);
    }
    public struct Abonnent
    {
        public string Id { get; set; }
        public string Address { get; set; }
        public string Name { get; set; }
        public string[] Subsribes { get; set; }
        
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

        public async Task SendMessageAsync(string message_body, string sender)
        {
            var configuration = new ConfigurationBuilder();

            try
            {
                foreach (var abo in Abonnents.Where(x => x.Subsribes.Contains(sender)))
                {
                    
                    var message = new MimeMessage();
                    message.From.Add(new MailboxAddress("Absender Name", "michael.schatzl@at.bosch.com"));
                    message.To.Add(new MailboxAddress("Empfänger Name", abo.Address));
                    message.Subject = "Test E-Mail aus .NET";
                    message.Body = new TextPart("html") { Text = "<b>Hallo!</b> Dies ist eine Test-Mail." };

                    using (var client = new SmtpClient())
                    {
                        // Verbindung zum SMTP-Server (z.B. Gmail, Outlook oder Firmenserver)
                       //var pass = _container["SMTP_Pass"];
                        await client.ConnectAsync("smtp.app.bosch.com", 587, MailKit.Security.SecureSocketOptions.Auto);
                        await client.AuthenticateAsync("HLS2HL@bosch.com", "RaKDRya5m3oHJ5Q5oAOORaKDRya5m3oHJ5Q5oAOO");
                        await client.SendAsync(message);
                        await client.DisconnectAsync(true);
                    }

                    //var abo = Abonnents.Where(m => m.Subsribes.Contains(sender));
                    //var client = new SmtpClient
                    //{
                    //    UseDefaultCredentials = true
                    //};
                    //foreach (var a in abo)
                    //{
                    //    var mail_message = new MailMessage(Application.Current.MainWindow.Name, a.Address);
                    //    mail_message.Subject = sender;
                    //    mail_message.Body = message;

                    //    client.Send(mail_message);
                    //} 
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
