using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

// add reference of OpenPop for email checking

using OpenPop.Mime;
using OpenPop.Pop3;

namespace AlertEmailParser
{
    class Program
    {
        
        public struct Alerts
        {
            public string facility;
            public string direction;
            public string segmentName;
            public double segmentID;
            public string type;
            public DateTime eventTime;
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
                    File.Create(m_exePath + "\\" + "AlertEmailParser_log_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt");

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
                    txtWriter.Write("{0} @ {1}", DateTime.Now.ToString("yyyy-MM-dd"), DateTime.Now.ToString("HH:mm:ss.fff") );
                    txtWriter.WriteLine(" :{0}", logMessage);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
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
                        if (unseenMessage.Headers.Subject.Contains("RITIS Speed Alert"))
                        {
                            string body1 = unseenMessage.MessagePart.MessageParts[0].GetBodyAsText(); //first body has the readable email 

                            string info = getBetween(body1, "Your alert ", ". ");
                            string ID = getBetween(info, "Exit ", "/");
                            alert.facility = getBetween(body1, "Your alert ", "-");
                            alert.direction = getBetween(info, "-", " from ");
                            alert.segmentName = getBetween(info, " from ", " has ");
                            double.TryParse( ID , out alert.segmentID);
                            alert.type = getBetween(body1, " has ", ". ");
                            alert.eventTime = unseenMessage.Headers.DateSent;

                            newAlerts.Add(alert); //save the alert

                            LogWriter.LogWrite("New alert: '" + info + "'. reported at " + alert.eventTime.ToString("yyyy-MM-dd HH:mm:ss"));
                            
                            client.DeleteMessage(i+1); // delete reviewed email

                        }
                        else if (unseenMessage.Headers.Subject.Contains("Frost"))
                        {
                            string body1 = unseenMessage.MessagePart.GetBodyAsText(); //first body has the readable email 

                            string info = getBetween(body1, "All:\r\n\r\n       ", "\n\r\n\r\n\r\n");
                            string ID = getBetween(info, "Exit ", "/");
                            alert.facility = getBetween(info, " ", "\r");
                            alert.direction = getBetween(info, "-", " from ");
                            alert.segmentName = getBetween(info, " from ", "\r");
                            double.TryParse(ID, out alert.segmentID);
                            alert.type = getBetween(body1, "(", ")");
                            alert.eventTime = unseenMessage.Headers.DateSent;

                            newAlerts.Add(alert); //save the alert

                            LogWriter.LogWrite("New alert: '" + getBetween(info, " ", "\r") + alert.type + "'. reported at " + alert.eventTime.ToString("yyyy-MM-dd HH:mm:ss"));

                        client.DeleteMessage(i+1); // delete reviewed email
                        }
                }
                LogWriter.LogWrite("Checked " + username + " email account...");
                LogWriter.LogWrite("Found " + newAlerts.Count + " new alerts.");
                // Return our new found messages
                return newAlerts;
            }
        }

        static void Main(string[] args)
        {
            Pop3Client EmailClient = new Pop3Client();

            string host = "outlook.office365.com", user = "innovationiac_amarillo@tti.tamu.edu", password = "Box46488";
            //string host = "outlook.office365.com", user = "innovationiac_bryan@tti.tamu.edu", password = "Boh74186"; 
            //string host = "outlook.office365.com", user = "innovationiac_atlanta@tti.tamu.edu", password = "Zop84471";
            int port = 995;
            bool useSsl = true;

            List<string> seenEmailID = new List<string>();

            try
            {
                List<Alerts> alerts = FetchUnseenAlerts(host, port, useSsl, user, password, seenEmailID);


                    
                
                

            }
            catch (Exception e) 
            {
                LogWriter.LogWrite("Cannot Connect to Email server.");
                LogWriter.LogWrite(e.ToString());
            }

        }
    }
}
