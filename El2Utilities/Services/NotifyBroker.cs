using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Windows.Devices.Sms;

namespace El2Core.Services
{
    public interface INotifyBroker
    {
        List<Abonnent> Abonnents { get; }
        Abonnent? GetAbonnentById(string id);
        void SendMessage(string message, string sender);
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

        public void SendMessage(string message, string sender)
        {
            var abo = Abonnents.Where(m => m.Subsribes.Contains(sender));
            var client = new SmtpClient
            {
                UseDefaultCredentials = true
            };
            foreach (var a in abo)
            {
                var mail_message = new MailMessage(Application.Current.MainWindow.Name, a.Address);
                mail_message.Subject = sender;
                mail_message.Body = message;

                client.Send(mail_message);
            }
        }
    }
}
