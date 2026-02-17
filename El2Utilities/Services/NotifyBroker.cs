using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using MimeKit;
using MailKit.Net.Smtp;

namespace El2Core.Services
{
    public interface INotifyBroker
    {
        List<Abonnent> Abonnents { get; }
        Abonnent? GetAbonnentById(string id);
        Task SendMessageAsync(string message_body, string sender);
        void AddAbonnent(Abonnent abonnent);
        void RemoveAbonnent(Abonnent abonnent);
    }
    public struct Abonnent
    {
        public string Address { get; set; }
        public string Name { get; set; }
        public string[] Subsribes { get; set; }

    }
    public class NotifyBroker : INotifyBroker
    {
        private readonly List<Abonnent> abonnents = [];
        public List<Abonnent> Abonnents { get { return abonnents; } }
        public Abonnent? GetAbonnentById(string id)
        {
            return Abonnents.FirstOrDefault(m => m.Name == id);
        }

        public void AddAbonnent(Abonnent abonnent)
        {
            Abonnents.Add(abonnent);
        }

        public void RemoveAbonnent(Abonnent abonnent)
        {
            Abonnents.Remove(abonnent);
        }

        public async Task SendMessageAsync(string message_body, string sender)
        {
  

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("Absender Name", "sender@beispiel.de"));
            message.To.Add(new MailboxAddress("Empfänger Name", "empfaenger@beispiel.de"));
            message.Subject = "Test E-Mail aus .NET";
            message.Body = new TextPart("html") { Text = "<b>Hallo!</b> Dies ist eine Test-Mail." };

            using (var client = new SmtpClient())
            {
                // Verbindung zum SMTP-Server (z.B. Gmail, Outlook oder Firmenserver)
                await client.ConnectAsync("smtp-mail.outlook365.de", 587, MailKit.Security.SecureSocketOptions.StartTls);
               // await client.AuthenticateAsync("benutzername", "passwort");
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
}
