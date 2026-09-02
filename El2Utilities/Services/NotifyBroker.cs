using El2Core.Models;
using El2Core.Utils;
using Microsoft.Extensions.Configuration;
using MimeKit;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

//using System;
//using System.Net;
//using System.Net.Mail;
//using System.Threading.Tasks;

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
    /// <summary>
    /// Represents a subscriber (Abonnent) with properties for ID, address, name, subscriptions, and workplaces.
    /// </summary>
    public struct Abonnent
    {
        public string Id { get; set; }
        public string Address { get; set; }
        public string Name { get; set; }
        public SubscribeType[] Subsribes { get; set; }
        public List<string> WorkPlaces { get; set; }
        
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
    /// <summary>
    /// Represents a notification broker that manages subscribers (Abonnents) and sends messages to them based on their subscriptions.
    /// </summary>
    public class NotifyBroker : INotifyBroker
    {
        private static IConfiguration? _container;
        public List<Abonnent> Abonnents { get; set; } = new List<Abonnent>();
        public bool IsChanged { get; set; }
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

        //public async Task SendEmailAsync(string toEmail, string subject, string body)
        //{
        //    // Annahme: SMTP-Server von Bosch. Passen Sie diesen bei Bedarf an.
        //    string smtpHost = "smtp.bosch.com";
        //    int smtpPort = 587; // Standard-Port für TLS

        //    // --- Platzhalter für RB-PAM-Authentifizierung ---
        //    // Hier sollten Sie Ihre Logik zum Abrufen der Anmeldeinformationen 
        //    // aus RB-PAM oder einem anderen sicheren Speicher (z.B. Azure Key Vault) einfügen.
        //    // Ersetzen Sie die folgenden Zeilen durch Ihren Code zum Abrufen von Benutzername und Passwort.
        //    string userName = "Ihr-Sysuser-Oder-Email";
        //    string password = "Ihr-Abgerufenes-Passwort";
        //    // --- Ende des Platzhalters ---

        //    var smtpClient = new SmtpClient(smtpHost)
        //    {
        //        Port = smtpPort,
        //        Credentials = new NetworkCredential(userName, password),
        //        EnableSsl = true, // Wichtig für eine sichere Verbindung
        //    };

        //    var mailMessage = new MailMessage
        //    {
        //        From = new MailAddress(userName),
        //        Subject = subject,
        //        Body = body,
        //        IsBodyHtml = true, // Auf 'true' setzen, wenn der Body HTML enthält
        //    };
        //    mailMessage.To.Add(toEmail);

        //    try
        //    {
        //        await smtpClient.SendMailAsync(mailMessage);
        //        Console.WriteLine("E-Mail erfolgreich versendet.");
        //    }
        //    catch (Exception ex)
        //    {
        //        // Gibt die Fehlermeldung aus, wenn das Senden fehlschlägt.
        //        Console.WriteLine($"Fehler beim Senden der E-Mail: {ex.Message}");
        //    }
        //}

        /// <summary>
        /// Sends a message to all subscribers (Abonnents) who have subscribed to the specified sender type. (SMTP)
        /// The message body is split into parts using the ASCII character 29 as a delimiter, and the message is sent in HTML format.
        /// </summary>
        /// <param name="message_body"></param>
        /// <param name="sender"></param>
        /// <returns></returns>
        public async Task SendMessageAsync(string message_body, SubscribeType sender)
        {
            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("HlP Lieferliste", "Lieferliste.HlP@at.bosch.com"));

            var mb = message_body?.Split((char)29) ?? Array.Empty<string>();
            var first = mb.Length > 0 ? mb[0] : string.Empty;
            var second = mb.Length > 1 ? mb[1] : string.Empty;
            var third = mb.Length > 2 ? mb[2] : string.Empty;
            var fourth = mb.Length > 3 ? mb[3] : string.Empty;

            // Add recipients who subscribed to this type
            foreach (var abo in Abonnents.Where(x => x.Subsribes != null && x.Subsribes.Contains(sender)))
            {
                if (abo.WorkPlaces.Contains(fourth) || abo.WorkPlaces.Count == 0)
                message.To.Add(new MailboxAddress(abo.Name, abo.Address));
            }

            if (message.To.Count == 0)
                return;

            message.Subject = "Abonnierte Nachricht";

            message.Body = new TextPart("html") { Text = $"<b>Hallo! Abonnents<p>Nachricht von {sender.Description()}</p></b><p>{first} hat folgendes geschrieben.</p>" +
                $"<p>{second}</p><p></p>Referenz: {third} - {fourth}" };

            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                await client.ConnectAsync("smtp.app.bosch.com", 587, MailKit.Security.SecureSocketOptions.Auto);
                client.Authenticate("HLS2HL@bosch.com", "d0K1nzTUj85x");
                await client.SendAsync(message);
                await client.DisconnectAsync(true);
            }
        }
        /// <summary>
        /// Updates an existing subscriber (Abonnent) in the list based on the provided value.
        /// If the subscriber with the same ID exists, it checks if the subscriptions and workplaces are the same. If they are different,
        /// it replaces the existing subscriber with the new value. Returns true if the update was successful, false otherwise.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool UpdateAbonnent(Abonnent value)
        {
            var idx = Abonnents.FindIndex(m => m.Id == value.Id);
            if (idx < 0)
                return false;

            var existing = Abonnents[idx];
            // If subscriptions are the same (reference or sequence), nothing to do
            if (existing.Subsribes == value.Subsribes || (existing.Subsribes != null && value.Subsribes != null && existing.Subsribes.SequenceEqual(value.Subsribes))
                && existing.WorkPlaces == value.WorkPlaces)
                return false;

            // Replace the item in the list (Abonnent is a struct)
            Abonnents[idx] = value;
            return true;
        }
    }
}
