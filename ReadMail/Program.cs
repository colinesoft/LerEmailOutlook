using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadMail
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = Emails.GetEmails(false, @"c:\temp\anexos");
            int i = 1;
            foreach (var mail in mails)
            {
                Console.WriteLine("Mail No: " + i);
                Console.WriteLine("Received from: " + mail.De);
                Console.WriteLine("Subject: " + mail.Assunto);
                //Console.WriteLine("Body: " + mail.Corpo);
                if(mail.Anexos.Count > 0)
                {
                    Console.WriteLine("ANEXOS:");
                    foreach (var anexo in mail.Anexos)
                        Console.WriteLine(anexo);
                }
                Console.WriteLine(" ");
                i++;
            }
        }
    }
}
