using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;

namespace ReadMail
{
    public class Emails
    {
        public string De { get; set; }
        public string Assunto { get; set; }
        public string Corpo { get; set; }
        public List<string> Anexos { get; set; }

        /// <summary>
        /// Lê os emails da caixa de entrada
        /// </summary>
        /// <returns></returns>
        public static List<Emails> GetEmails(bool onlyUnread = false, string saveInFolder = "")
        {
            Application outlook = null;     //OUTLOOK
            MAPIFolder inbox = null;        //CAIXA DE ENTRADA
            Items mailItems = null;         //ITENS LIDOS NA CAIXA DE ENTRADA

            //RETORNO
            List<Emails> emailsRetorno = new List<Emails>();

            //Os emails válidos serão atribuído para essa classe mailLido
            MailItem mailLido;

            try
            {
                //HABILITA O OUTLOOK 
                outlook = new Application();
                //OBTEM A CAIXA DE ENTRADA DO OUTLOOK
                inbox = outlook.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

                //Carrega os itens lidos da caixa de entrada.
                //OBS: Nem tudo que foi lido é EMAIL de fato
                // Verifica também se deve ser apenas emails ainda não lidos
                mailItems = onlyUnread ? inbox.Items.Restrict("[Unread]=true") : inbox.Items;
                                
                foreach (Object item in mailItems)
                {
                    //Testa se o item lido da cx de entrada é um MailItem
                    if ((item as MailItem) != null)
                    {
                        //FAz um CAST para facilitar o uso
                        mailLido = (item as MailItem);

                        //faz um match do email lido para o email que será retornado
                        var mailRetorno = new Emails();

                        mailRetorno.De = mailLido.SenderEmailAddress;
                        mailRetorno.Assunto = mailLido.Subject;
                        mailRetorno.Corpo = mailLido.Body;
                        mailRetorno.Anexos = new List<string>();

                        //Verifica se existe algum anexo atrelado ao email
                        //Os anexos começar pelo indice 1
                        var qtdeAnexo = mailLido.Attachments.Count;
                        for (int i = 1; i <= qtdeAnexo; i++)
                        {
                            var fileName = mailLido.Attachments[i].FileName;
                            mailRetorno.Anexos.Add(fileName);
                            var extensao = Path.GetExtension(fileName);

                            if (!string.IsNullOrEmpty(saveInFolder) && extensao.ToLower() == ".xlsx")
                                mailLido.Attachments[i].SaveAsFile(Path.Combine(saveInFolder, fileName));
                        }
                        emailsRetorno.Add(mailRetorno);
                        ReleaseComObject(mailLido);
                    }
                }
            } 
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Limpa da memória os emails lidos
                ReleaseComObject(mailItems);
                ReleaseComObject(inbox);
                ReleaseComObject(outlook);
            }
            return emailsRetorno;
        }

        //Limpa obj da memória
        private static void ReleaseComObject(object obj)
        {
            if(obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
