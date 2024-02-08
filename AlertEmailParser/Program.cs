using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

// add reference of OpenPop for email checking

using OpenPop.Mime;
using OpenPop.Pop3;
using static AlertEmailParser.Program;

//add support for sending emails
using System.Net;
using System.Net.Mail;
using System.Reflection.Metadata;

namespace AlertEmailParser
{



    
    class Program
    {


        public struct Alerts
        {
            public string facility;
            public string segmentName;
            public double segmentID;
            public string type;
            public DateTime eventTime;
            public double life;
        }

        public struct FrostSolutionsSite
        {
            public string facility;
            public bool isBelow;
            public double temp;
            public double tempLife;
            public bool isRoadCondition;
            public string type;
            public double roadLife;
            public DateTime lastUpdatedCondition;
            public DateTime lastUpdatedTemp;
            public bool updateEmailSent;
        }

        public static string getBetween(string strSource, string strStart, string strEnd)
        {
            if (strSource.Contains(strStart) && strSource.Contains(strEnd))
            {
                int Start, End;
                Start = strSource.IndexOf(strStart, 0) + strStart.Length;
                End = strSource.IndexOf(strEnd, Start);
                return strSource.Substring(Start, End - Start);
            }

            return "";
        }
        public static class LogWriter
        {
            private static string m_exePath = string.Empty;
            public static void LogWrite(string logMessage)
            {
                m_exePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!File.Exists(m_exePath + "\\" + "AlertEmailParser_log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                {
                    File.Create(m_exePath + "\\" + "AlertEmailParser_log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt").Close();
                }

                try
                {
                    using (StreamWriter w = File.AppendText(m_exePath + "\\" + "AlertEmailParser_log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt"))
                    {
                        AppendLog(logMessage, w);
                        w.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }

            private static void AppendLog(string logMessage, TextWriter txtWriter)
            {
                try
                {
                    txtWriter.Write("UTC Time: {0} @ {1}", DateTime.UtcNow.ToString("yyyy-MM-dd"), DateTime.UtcNow.ToString("HH:mm:ss.fff") );
                    txtWriter.Write(",CST Time: {0} @ {1}", DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("HH:mm:ss.fff"));
                    txtWriter.WriteLine(" :{0}", logMessage);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            
        }
        public static void SendAlertEmailMessage(SmtpClient client, NetworkCredential credential, FrostSolutionsSite site)
            {
                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {

                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = credential;
                    smtp.EnableSsl = true;

                    MailAddress from = new MailAddress("innovationiacamarillo@gmail.com");
                    MailAddress to = new MailAddress("d-florence@tti.tamu.edu");
                    
                    MailMessage email = new MailMessage(from, to);

                    email.To.Add(new MailAddress("k-balke@tti.tamu.edu"));

                    email.ReplyToList.Add(new MailAddress("d-florence@tti.tamu.edu"));



                    //set body-message and subject
                    if (site.updateEmailSent)
                    {
                        email.Subject = "PCMS Weather Alert Message Posted - " + site.facility;
                        email.Body = "<br><em>Please note: This is only a test Email.</em><br><br>Weather alert for PCMS " + site.facility + " <em><b>activated.</em></b><br><br>";
                    }
                    else
                    {
                        email.Subject = "PCMS Weather Alert Message Deactivated - " + site.facility;
                        email.Body = "<br><em>Please note: This is only a test Email.</em><br><br>Weather alert for PCMS " + site.facility + " <em><b>deactivated.</em></b><br><br>";
                    }

                    email.SubjectEncoding = System.Text.Encoding.UTF8;
                    email.BodyEncoding = System.Text.Encoding.UTF8;

                    //text or html
                    email.IsBodyHtml = true;

                    smtp.Send(email);
                }
            }
        /***********    From OpenPop.net examples (IMAP protocol  ***************/
        /// <summary>
        /// Example showing:
        ///  - how to use UID's (unique ID's) of messages from the POP3 server
        ///  - how to download messages not seen before
        ///    (notice that the POP3 protocol cannot see if a message has been read on the server
        ///     before. Therefore the client need to maintain this state for itself)
        /// </summary>
        /// <param name="hostname">Hostname of the server. For example: pop3.live.com</param>
        /// <param name="port">Host port to connect to. Normally: 110 for plain POP3, 995 for SSL POP3</param>
        /// <param name="useSsl">Whether or not to use SSL to connect to server</param>
        /// <param name="username">Username of the user on the server</param>
        /// <param name="password">Password of the user on the server</param>
        /// <param name="seenUids">
        /// List of UID's of all messages seen before.
        /// New message UID's will be added to the list.
        /// Consider using a HashSet if you are using >= 3.5 .NET
        /// </param>
        /// <returns>A List of new Messages on the server</returns>
        public static List<Alerts> FetchUnseenAlerts(string hostname, int port, bool useSsl, string username, string password, List<string> seenUids)
        {
            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                    client.Connect(hostname, port, useSsl);

                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                int messageCount = client.GetMessageCount();
                LogWriter.LogWrite("found " + messageCount + " emails.");
                List<Message> allMessages = new List<Message>(messageCount);

                // Fetch all the current uids seen
                List<string> uids = client.GetMessageUids();

                // Create a list we can return with all new alerts
                List<Alerts> newAlerts = new List<Alerts>();
                // All the messages not seen by the POP3 client
                for (int i = 0; i < uids.Count; i++)
                {
                        // the uids list is in messageNumber order - meaning that the first
                        // uid in the list has messageNumber of 1, and the second has 
                        // messageNumber 2. Therefore we can fetch the message using
                        // i + 1 since messageNumber should be in range [1, messageCount]
                        OpenPop.Mime.Message unseenMessage = client.GetMessage(i + 1);
                        Alerts alert = new Alerts();
                        // read messages for alerts
                       if (unseenMessage.Headers.Subject.Contains("Frost") && unseenMessage.Headers.From.Address.Contains("info@frostsolutions.io") )
                        {
                            string body1 = unseenMessage.MessagePart.GetBodyAsText(); //first body has the readable email 

                            string info = getBetween(body1, "All:\r\n\r\n       ", "\r\n\r\n");
                            string ID = getBetween(info, "Exit ", "/");
                            alert.facility = getBetween(info, " ", "\r");
                            alert.segmentName = getBetween(info, "— ", "\r");
                            double.TryParse(getBetween(body1, "trigger again until ", " hours have passed."), out alert.life);
                            double.TryParse(ID, out alert.segmentID);
                            alert.type = getBetween(body1, "(", ")");

                            alert.eventTime = unseenMessage.Headers.DateSent;

                        

                            

                            LogWriter.LogWrite("New alert at: \"" + alert.facility + "\". Alert type: \"" + alert.type + "\". Segment: \"" + alert.segmentName + "reported at: " + alert.eventTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            newAlerts.Add(alert);
                        
                        }
                }
                LogWriter.LogWrite("Checked " + username + " email account...");
                LogWriter.LogWrite("Found " + newAlerts.Count + " new alerts.");
                // Return our new found messages
                return newAlerts;
            }
        }

        public static List<FrostSolutionsSite> ManageAlerts(List<Alerts> activeAlerts, List<FrostSolutionsSite> sites) //check for timeout of alerts based on the life
        {
            int numAlerts = activeAlerts.Count;
            

            for (int i =  0; i < numAlerts; i++)
            {
                FrostSolutionsSite site = sites.Find(FrostSolutionSite => FrostSolutionSite.facility.Equals(activeAlerts[i].facility));

                if (site.facility == null)
                {
                    site.facility = activeAlerts[i].facility;
                    sites.Add(site);
                }
            }

            int numsites = sites.Count;
            TimeSpan buffer = TimeSpan.FromMinutes(15);

            for (int j = 0; j < numsites; j++)
            {
                FrostSolutionsSite site = sites[j];
                for (int i = 0; i < numAlerts; i++)
                {
                    Alerts alert = activeAlerts[i];
                    
                    if (alert.facility == sites[j].facility)
                    {
                        if (alert.type.Contains("Road"))
                        {
                            site.isRoadCondition = true;
                            site.roadLife = alert.life;
                            site.lastUpdatedCondition = alert.eventTime;
                        }
                        else if (alert.type.Contains(">"))
                        {
                            site.isBelow = false;
                            site.tempLife = alert.life;
                            double.TryParse(getBetween(alert.type, "Surface Temp > ", "°"), out site.temp);
                            site.lastUpdatedTemp = alert.eventTime;

                        }
                        else if (alert.type.Contains("<"))
                        {
                            site.isBelow = true;
                            site.tempLife = alert.life;
                            double.TryParse(getBetween(alert.type, "Surface Temp < ", "°"), out site.temp);
                            site.lastUpdatedTemp = alert.eventTime;
                        }
                    }

                }

                if( DateTime.UtcNow - site.lastUpdatedTemp.AddHours(site.tempLife) > buffer)
                {
                    site.isBelow = false;
                    site.tempLife = 99;
                    site.lastUpdatedTemp = DateTime.UtcNow;
                }

                if( DateTime.UtcNow - site.lastUpdatedCondition.AddHours(site.roadLife) > buffer)
                {
                    site.isRoadCondition = false;
                    site.roadLife = 99;
                    site.lastUpdatedCondition= DateTime.UtcNow;
                }

                sites[j] = site;
                LogWriter.LogWrite("Updated: " + site.facility + " Temperature Data Below: "  + site.isBelow + " Temp Threshold: "+ site.temp + " Road Condition: " + site.isRoadCondition);
            }
            return sites;
        }

        public static void DMSAlert(List<FrostSolutionsSite> sites, PCMSUpdater.PCMS.SignSpecs sign, SmtpClient client, NetworkCredential credential)
        {
            int numsites = sites.Count;
            string messageToPost;
            string PCMSresponse = string.Empty;

            for (int i = 0; i < numsites; i++)
            {
                FrostSolutionsSite site = sites[i];

                if (site.isBelow && site.isRoadCondition)
                {
                    //send information to DMS sign
                    
                    LogWriter.LogWrite("Sending message to PCMS for " + site.facility + "(Function still incomplete)");
                    Console.Write("...\tPosting weather warning message!");

                    messageToPost = sign.alertMessage;
                    if (site.updateEmailSent == false)
                    {
                        site.updateEmailSent = true;
                        SendAlertEmailMessage(client, credential, site);
                    }
                }
                else
                {
                    messageToPost = sign.defaultMessage;
                    if (site.updateEmailSent == true)
                    {
                        site.updateEmailSent = false;
                        SendAlertEmailMessage(client, credential, site);
                    }
                    
                }

                //The Sign updater
                //Not ready to deploy yet!
                //PCMSresponse = PCMSUpdater.PCMS.UpdateMessage(sign, messageToPost);
                //LogWriter.LogWrite("Response form PCMSUpdater: " + PCMSresponse);
            }
        }

        public static void CheckWeatherCondition(SmtpClient client, NetworkCredential credential, string host, string user, string password, int port, bool useSsl, List<string> seenEmailID, List<Alerts> alerts, List<FrostSolutionsSite> SupportedSites, PCMSUpdater.PCMS.SignSpecs sign)
        {
            Console.Write("\nSystem Checking at " + DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                alerts = FetchUnseenAlerts(host, port, useSsl, user, password, seenEmailID);

            }
            catch (Exception e)
            {
                LogWriter.LogWrite("Failed to check email inbox.");
                LogWriter.LogWrite(e.ToString());
            }

            SupportedSites = ManageAlerts(alerts, SupportedSites);

            DMSAlert(SupportedSites, sign, client, credential);

        }

        static async Task Main()
        {
            Pop3Client EmailClient = new Pop3Client();

            string host = "pop.gmail.com", user = "innovationiacamarillo@gmail.com", password = "mncq zsxw zbhh xtwz", smtpPassword = "qtlj aqds vhgt rfjp";
            //string host = "outlook.office365.com", user = "innovationiac_amarillo@tti.tamu.edu", password = "Box46488";

            int port = 995;
            bool useSsl = true;

            List<string> seenEmailID = new List<string>();
            List<Alerts> alerts = new List<Alerts>();
            List<FrostSolutionsSite> SupportedSites = new List<FrostSolutionsSite>();

            Console.WriteLine("Starting Weather Monitoring of : " + user);

            Console.WriteLine("\n...Running!");
            Console.WriteLine("\n\nUpdating sign every 5 minutes...\n\n");
            
            //var timer = new PeriodicTimer(TimeSpan.FromSeconds(10));  //Use a 10 second gap for debugging
            var timer = new PeriodicTimer(TimeSpan.FromMinutes(5));

            SmtpClient mySmtpClient = new SmtpClient("smtp.gmail.com", 587);
            mySmtpClient.UseDefaultCredentials = false;
            NetworkCredential basicAuthinticationInfo = new NetworkCredential(user, smtpPassword);
            

            

            PCMSUpdater.PCMS.SignSpecs sign = new PCMSUpdater.PCMS.SignSpecs();
            sign.id = "EB-PCMS4";
            sign.IPAddress = "166.239.126.13";
            sign.community = "Public";
            sign.description = "Amarillo I-40 PCMS4";
            sign.enableUpdates = true;
            sign.image = "http://";
            sign.latitute = 31.99962031927457;
            sign.longitute = -101.9262841747892;
            sign.slot = ".3.100";
            sign.defaultMessage = "[fo6][jp3][jl3]DRIVE[nl]SAFELY";
            sign.alertMessage = "[pt20o0][jl3]WATCH[nl]FOR[nl]ICE[np][pt20o0][jl3][fo1]ON[nl]ROADWAY[nl]AHEAD";

            //FrostSolutionsSite site = new FrostSolutionsSite();
            //site.facility = "West of Adrian";
            //SendAlertEmailMessage(mySmtpClient, basicAuthinticationInfo, site);
            

            CheckWeatherCondition(mySmtpClient, basicAuthinticationInfo, host, user, password, port, useSsl, seenEmailID, alerts, SupportedSites, sign);

            while (await timer.WaitForNextTickAsync())
            {
                CheckWeatherCondition(mySmtpClient, basicAuthinticationInfo,host, user, password, port, useSsl, seenEmailID, alerts, SupportedSites, sign);
            }
        }
    }
}
