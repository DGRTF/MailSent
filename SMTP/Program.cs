using System;
using System.Net.Mail;
using System.Net;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Security;
using ADODB;

namespace SMTP
{
    class Program
    {





        interface IFileString
        {
            List<string> MailAdd { get; }

        }







        class Exls : IFileString
        {
            public Exls()
            {
                MailAdd = Connect();
            }



            public List<string> MailAdd { get; }
            public string Path { get; private set; }



            public List<string> Connect()
            {

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                MessageBox.Show($"Click 'OK' and specify the address file");
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Path = openFileDialog1.FileName;
                    }
                    catch (SecurityException ex)
                    {
                        MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                        $"Details:\n\n{ex.StackTrace}");
                    }
                }
                //Console.WriteLine(Path);
                //Console.Read();
                string con = "Data Source=";
                string con1 = "; Mode=Read;Extended Properties='Excel 12.0'";
                List<string> collection = new List<string>();
                var v = new ADODB.Connection
                {
                    //ConnectionString = @"Data Source=C:\Users\Public\Desktop\ItCompany.xlsx; Mode=Read;Extended Properties='Excel 12.0'",
                    ConnectionString = con + Path + con1,
                    Provider = "Microsoft.ACE.OLEDB.12.0",
                    Mode = ConnectModeEnum.adModeRead
                };
                v.Open();
                var vv = new Recordset();
                //{
                //    CursorType = (CursorTypeEnum)1
                //};
                vv.Open(ActiveConnection: v.ConnectionString, Source: "SELECT * FROM [Лист1$]");
                int i = 0;

                while (vv.EOF == false)
                {

                    foreach (Field go in vv.Fields)
                    {

                        if (go.Name == "почта")
                        {
                            Console.WriteLine(go.Value);
                            Object t = (object)go.Value;
                            collection.Add(t.ToString());
                            i++;
                        }
                    }
                    vv.MoveNext();
                }
                Console.WriteLine("All Records " + i);
                v.Close();
                vv.Close();
                return collection;
            }
        }








        class SMTPclientt
        {
            public string MailFrom { get; set; }
            public IFileString ToMail { get; set; }
            //public string Path { get; set; }
            //public string UserName { get; set; }
            public string UserPass { get; set; }
            public string Sub { get; set; } = "";
            public string Bo { get; set; } = "";
            public async void SMTPSendGmail()
            {

                using (SmtpClient client = new SmtpClient("smtp.gmail.com", 587)
                {
                    EnableSsl = true,
                    UseDefaultCredentials = true,
                    Credentials = new NetworkCredential(MailFrom, UserPass)
                })
                {

                    MailMessage mess = new MailMessage();
                    //{
                    mess.Body = Bo;
                        mess.From = new MailAddress(MailFrom);
                        mess.Subject = Sub;
                    //};
                    //mess.Attachments.Add(new Attachment(Path));
                    foreach (string to in ToMail.MailAdd)
                    {
                        if (to != "")
                        {
                            mess.To.Add(new MailAddress(to));
                            await client.SendMailAsync(mess);
                            mess.To.Clear();
                        }
                    }
                    mess.Dispose();
                }
            }
       
        }





        [STAThread]
        static void Main()
        {
            SMTPclientt client = new SMTPclientt
            {
                //MailFrom = "Gena",
                ToMail = new Exls()
            };
            Console.WriteLine("Please, enter Username");
            client.MailFrom = Console.ReadLine();
            Console.WriteLine("Please, enter Password");
            client.UserPass = Console.ReadLine();
            Console.WriteLine("Please, enter Subject");
            client.Sub = Console.ReadLine();
            Console.WriteLine("Please, enter Text Message");
            client.Bo = Console.ReadLine();





            client.SMTPSendGmail();
            Console.ReadKey();
        }
    }
}
