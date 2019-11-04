using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace DiskSpace
{
    public partial class Service1 : ServiceBase
    {
        Timer timer;
        private float acceptablePercentage = 0.10f;
        private string[] toEmails = { };
        

        public Service1()
        {
            InitializeComponent();
            var minuteInMiliseconds = 60000;
            timer = new Timer(minuteInMiliseconds);
            timer.Elapsed += delegate { onTick(); };
        }

        private void onTick()
        {
            var gigabyteInBytes = 1024 * 1024 * 1024;
            var lowOnSpace = false;
            foreach (var drive in DriveInfo.GetDrives())
                if (drive.IsReady)
                    if ((float)drive.AvailableFreeSpace / drive.TotalSize <= acceptablePercentage && (float)drive.TotalSize / gigabyteInBytes > 10f)
                        lowOnSpace = true;

            var message = "One of the drives has under " + acceptablePercentage * 100 + "% free space. Here is the current status of all: \n\n";
            if (lowOnSpace)
            {
                foreach (var drive in DriveInfo.GetDrives())
                    if (drive.IsReady)
                    {
                        message += drive.Name + " Available space: " + (float)drive.AvailableFreeSpace / (1024 * 1024 * 1024) + "/" + (float)drive.TotalSize / (1024 * 1024 * 1024) + " GB";
                        if ((float)drive.AvailableFreeSpace / drive.TotalSize <= acceptablePercentage && (float)drive.TotalSize / gigabyteInBytes > 10f)
                            message += "  <------";
                        message += "\n";
                    }

                SendEmail(String.Join(",", toEmails), "", "", "[" + Environment.MachineName +  "] Low on disk space", message);
            }
        }

        public void Start()
        {
            OnStart(null);
            onTick();
        }
        protected override void OnStart(string[] args)
        {
            timer.Start();
        }

        protected override void OnStop()
        {
        }

        public static void SendEmail(String ToEmail, string cc, string bcc, String Subj, string Message)
        {
            //Reading sender Email credential from web.config file  

            string HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
            string FromEmailid = ConfigurationManager.AppSettings["FromMail"].ToString();
            string Pass = ConfigurationManager.AppSettings["Password"].ToString();

            //creating the object of MailMessage  
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(FromEmailid); //From Email Id  
            mailMessage.Subject = Subj; //Subject of Email  
            mailMessage.Body = Message; //body or message of Email  
            mailMessage.IsBodyHtml = false;

            string[] ToMuliId = ToEmail.Split(',');
            foreach (string ToEMailId in ToMuliId)
            {
                mailMessage.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
            }


            string[] CCId = cc.Split(',');

            foreach (string CCEmail in CCId)
            {
                if (CCEmail != string.Empty)
                    mailMessage.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
            }

            string[] bccid = bcc.Split(',');

            foreach (string bccEmailId in bccid)
            {
                if (bccEmailId != string.Empty)
                    mailMessage.Bcc.Add(new MailAddress(bccEmailId)); //Adding Multiple BCC email Id  
            }
            SmtpClient smtp = new SmtpClient();  // creating object of smptpclient  
            smtp.Host = HostAdd;              //host of emailaddress for example smtp.gmail.com etc  

            //network and security related credentials  

            smtp.EnableSsl = true;
            NetworkCredential NetworkCred = new NetworkCredential();
            NetworkCred.UserName = mailMessage.From.Address;
            NetworkCred.Password = Pass;
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mailMessage); //sending Email  
        }
    }
}
