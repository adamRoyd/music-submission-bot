using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ID3Bot
{
    class MailService
    {
        private readonly IConfiguration _configuration;

        public MailService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public async Task SendMailAsync()
        {
            Console.WriteLine($"Sending mail");

            try
            {
                using (var message = new MailMessage())
                {
                    message.To.Add(new MailAddress("adamboothroyd1@gmail.com", "Adam"));
                    message.From = new MailAddress("hardfactmusic@gmail.com", "Hard Fact");
                    message.Subject = "ID3 Tag Processing results";
                    message.Body = GetBody();
                    message.IsBodyHtml = true;

                    using (var client = new SmtpClient("smtp.gmail.com"))
                    {
                        client.Port = 587;
                        client.Credentials = new NetworkCredential("hardfactmusic@gmail.com", _configuration["EMAIL_PASSWORD"]);
                        client.EnableSsl = true;
                        await client.SendMailAsync(message);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error sending mail. Exception {e}");
            }

        }

        private string GetBody()
        {
            var template = File.ReadAllText("EmailTemplate.html");

            return template;
        }
    }
}
