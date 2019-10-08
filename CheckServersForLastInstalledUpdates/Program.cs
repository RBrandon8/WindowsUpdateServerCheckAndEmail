using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Microsoft.Win32;
using System.DirectoryServices;
using System.Collections;
using System.Net;
using System.Net.Mail;
using System.IO;
using WUApiLib;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;
using CheckServersForLastInstalledUpdates.Classes;


namespace CheckServersForLastInstalledUpdates
{
    class Program
    {
        //Sets Debug and Error Log directory and number of servers
        //to test with debug. 
        static bool debug = false;
        static string output = "";
        static string debugDir = @"C:\temp\";
        static string debugFilePath = debugDir + "UpdateDebug.txt";
        static int debugNumTestServers = 20;

        //Required email variables for your environment
        static readonly string emailHost = "your_email_server"; //smtp.gmail.com 
        static readonly string emailFrom = "NoReply@yourdomain.com"; //
        static readonly string emailRecpient = "youremail@yourdomain.com";
        static readonly int emailPort = 25; //use 587 for Gmail requires auth see SendMailWithTable method



        //Required Active Directory variables and WUApi calls
        static readonly string serverOU = "ou=Servers,dc=domain;dc=com";
        static readonly List<string> excludedServers = new List<string>() {}; //server names that shouldn't be checked.
        static readonly int waitTimeForServerCall = 300000; //5 minutes

        //Email Subject Formatting variables
        static readonly string alternatingRowColor = "#f2f2f2";  //Not working currently
        static readonly string headerColor = "#0072C6";
        static readonly string tableHeaderColor = "#0072C6";
        static readonly string tableHeaderFontColor = "#FFFFFF";

        static void Main(string[] args)
        {
            try
            {
                //Check if any arguments are passed.  Looking for debug command
                //to add additional logging to troubleshoot server issues.
                if (args.Count() >= 1)
                {
                    foreach (string arg in args)
                        if (arg.ToUpper().Trim() == "DEBUG")
                            debug = true;
                }

                //Retrive list of servers with updates and send HTML formated email
                List<Server> servers = GetServerList();
                /*    //Creates a test list and sorts for testing and formatting purposes.
                        new List<Server>() { new Server("Test1", DateTime.Now, 10, 5),
                        new Server("Test1", DateTime.Now, 10, 5),new Server("Test6", DateTime.Now, 12, 5)
                        ,new Server("Test2", DateTime.Now, 0, 0),new Server("Test7", DateTime.Now, 3, 5)
                        ,new Server("Test3", DateTime.Now, 20, 15),new Server("Test8", DateTime.Now, 2, 5)
                        ,new Server("Test4", DateTime.Now, 0, 5),new Server("Test9", DateTime.Now, 1, 5)
                        ,new Server("Test5", DateTime.Now, 5, 5),new Server("Test10","C8000710: Disk is Full")};
                var myList = servers.ToList();
                myList.Sort(delegate (Server c1, Server c2) { return c2.NumImpUpdates.CompareTo(c1.NumImpUpdates); });
                */

                SendMailwithTable("Server Last Update Listing", emailFrom, emailRecpient, servers);
                    //myList); //For Testing
            }
            catch (Exception e)
            {
                WriteErrorLog(e.Message);
            }
        }

        /*
         *   Retrieves all computer objects in an Active
         *   Directory OU passed from initialization variable.
         *   Loops through objects and then calls WUApi to 
         *   retrieve number of updates.
         */
        private static List<Server> GetServerList()
        {

            // Retrieve OU objects from AD and store in a List
            using (DirectoryEntry de = new
                DirectoryEntry("LDAP://" + serverOU))
            {
                DirectorySearcher src = new DirectorySearcher();
                src.SearchRoot = de;
                src.SearchScope = System.DirectoryServices.SearchScope.Subtree;
                src.Filter = "(objectCategory=computer)";
                List<Server> list = new List<Server>();


                SearchResultCollection res = src.FindAll();
                List<string> servers = new List<string>();
                int cnt = 0;

                //Loop through checking for debug and if the server
                //is excluded.  Then save to a collection for WUApi 
                //call.
                foreach (SearchResult sc in res)
                {
                    foreach (string myCollection in sc.Properties["cn"])
                    {
                        if (cnt > debugNumTestServers && debug)
                            break;
                        else cnt++;

                        if (!excludedServers.Contains(myCollection.ToUpper()))
                            servers.Add(myCollection);
                        
                        //Add server name to debug file.  Can be used to compare objects processed
                        //with WUApi for long running or never returning computer objects.
                        if (debug)
                            File.AppendAllText(debugFilePath, myCollection + System.Environment.NewLine);
                    }
                }

                //Add a line break to differentiate between returned AD objects and objects
                //processed by WUApi.
                if (debug)
                    File.AppendAllText(debugFilePath, "~" + System.Environment.NewLine);

                //Create Cancellation token to prepare parallel for each run.
                CancellationTokenSource cts = new CancellationTokenSource();
                ParallelOptions po = new ParallelOptions();
                po.CancellationToken = cts.Token;
                po.MaxDegreeOfParallelism = System.Environment.ProcessorCount;

                //Add threading for long running operation.  Had to add an additional Task
                //inside each thread to allow the task to timeout if the server object 
                //never returns.
                try
                {
                    Parallel.ForEach(servers, po, (s) =>
                    {
                        Task t = Task.Run(() =>
                            {
                                try
                                {
                                    //Call Methods to get Windows Update info for server.
                                    DateTime lastInstall = GetLastWindowsUpdatesInstalledDate(s);
                                    Tuple<int, int> updates = GetServerUpdates(s);

                                    //Add to list and update console to show progress.
                                    Console.WriteLine("{0}: {1} {2},{3}", s, lastInstall.ToShortDateString(), updates.Item1, updates.Item2);
                                    list.Add(new Server(s, lastInstall, updates.Item1, updates.Item2));

                                    //Record that the server successfully processed for debugging
                                    if (debug)
                                        File.AppendAllText(debugFilePath, s + System.Environment.NewLine);
                                }
                                // Record the error message returned by WUApi call
                                catch (Exception e)
                                {
                                    String result = e.HResult.ToString("X");

                                    switch (result)
                                    {
                                        case "800706BA":
                                            result = result + ": The RPC server is unavailable";
                                            break;
                                        case "80070005":
                                            result = result + ": Access is denied";
                                            break;
                                        case "C8000710":
                                            result = result + ": Disk is Full";
                                            break;
                                        case "80072EFD":
                                            result = result + ": ERROR_WINHTTP_CANNOT_CONNECT or ERROR_INTERNET_CANNOT_CONNECT The attempt to connect to the server failed. / Internet cannot connect";
                                            break;
                                        case "8024001E":
                                            result = result + ": Operation did not complete because the service or system was being shut down.";
                                            break;
                                        default:
                                            //result initilization will return the HResult if switch doesn't match.
                                            //Record that the server errored for debugging.
                                            if (debug)
                                                File.AppendAllText(debugFilePath, s + System.Environment.NewLine); //+ ": " + result
                                            break;
                                    }

                                    //Add processed servers to list.
                                    Console.WriteLine(s + ": " + result);
                                    list.Add(new Server(s, result));
                                }
                            });
                        //Add a timed out server to the output list.  Servers that
                        //timeout mean that the waitTimeForServerCall timeout needs
                        //to be increased or the server is having issues.
                        TimeSpan ts = TimeSpan.FromMilliseconds(waitTimeForServerCall);
                        if (!t.Wait(ts))
                        {
                            Console.WriteLine(s + " Timed out");
                            list.Add(new Server(s, "Timed out, manually check server"));
                        }
                    });
                }
                //If cancellation token is raised stop processing.  Not in use currently
                //as no user input is recorded
                catch (OperationCanceledException e)
                {
                    File.AppendAllText(debugFilePath, e.Message + System.Environment.NewLine);
                }
                //Dispose cancellation token after completion.
                finally
                {
                    cts.Dispose();
                }

                //Add servers to output list to include them in email.
                foreach (string s in excludedServers)
                {
                    list.Add(new Server(s, "Server was Excluded"));
                }

                //Convert to list to sort by number of Important Updates found.
                var myList = list.ToList();
                myList.Sort(delegate (Server c1, Server c2) { return c2.NumImpUpdates.CompareTo(c1.NumImpUpdates); });

                return myList;
            }
        }

        /*
         *  Use WUApi to retrieve the update with the latest successful install date 
         */
        public static DateTime GetLastWindowsUpdatesInstalledDate(String serverName)
        {
            // Create WUApi instance and intialize return variable
            Type t = Type.GetTypeFromProgID("Microsoft.Update.Session", serverName);
            DateTime dt = DateTime.MinValue;

            try
            {
                //Call Update interface and search to get total number of updates installed
                //Then call interface to retrieve all updates.
                UpdateSession session = (UpdateSession)Activator.CreateInstance(t);
                IUpdateSearcher updateSearcher = session.CreateUpdateSearcher();
                int count = updateSearcher.GetTotalHistoryCount();
                IUpdateHistoryEntryCollection history = updateSearcher.QueryHistory(0, count);

                //Loop through and overwrite return date variable if install date is later.
                foreach (IUpdateHistoryEntry he in history)
                {
                    try
                    {
                        if (he.ResultCode == OperationResultCode.orcSucceeded && dt < he.Date)
                        {
                            dt = he.Date;
                        }

                    }
                    catch (Exception ex) {
                        WriteErrorLog(serverName + ":" + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(serverName + ":" + ex.Message);
            }

            return dt;
        }

        /*
         *  Call WUApi to get available updates.  Returns two int's 
         *  <Important Updates,Recommended Updates>
         */
        static public Tuple<int, int> GetServerUpdates(string serverName)
        {

            UpdateCollection uc = null;
            int cntImpUpdate = 0;
            int cntRecUpdate = 0;

            try
            {
                //Create WUApi session and filter results to only include Software
                //and exclude installed and hidden updates.
                Type t = Type.GetTypeFromProgID("Microsoft.Update.Session", serverName);
                UpdateSession session = (UpdateSession)Activator.CreateInstance(t);
                session.ClientApplicationID = "BR Server Update";
                IUpdateSearcher updateSearcher = session.CreateUpdateSearcher();
                ISearchResult iSResult = updateSearcher.Search("Type = 'Software' and IsHidden = 0 and IsInstalled = 0");
                uc = iSResult.Updates;

                //loop through updates and identify update priority.
                foreach (IUpdate u in uc)
                {
                    string cat = u.Categories[0].Name;
                    if (cat != "Tools" && cat != "Feature Packs" && cat != "Updates")
                        cntImpUpdate++;
                    else
                        cntRecUpdate++;
                }

            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }            

            return Tuple.Create(cntImpUpdate, cntRecUpdate);

        }

        /*
         * Creates the html return string for subject of email.
         */ 
        public static string FormatTableForMailSubject(List<Server> servers)
        {
            int cntImp = 0;
            int cntRec = 0;

            //Get a total count of all servers available updates.
            foreach (Server s in servers)
            {
                cntImp += s.NumImpUpdates;
                cntRec += s.NumRecUpdates;
            }



            string textBody =
                // HTML Style
                "<html><head><style>" +
                    "table {" +
                        "border - collapse: collapse;" +
                        "width: 800px;" +
                    "}" +
                    "th,tr {" +
                    "padding: 8px;" +
                    "}" +
                    "td{" +
                       "padding - right: 10px;" +
                    "}" +
                    "tr: nth - child(even) { background - color: '" + alternatingRowColor +"';}" +
                    "#errorStyle{padding-left: 10px;}" +
                    "</style></head>" +
                //Header formatting 
                "<header>" +
                "<h1><font color='" + headerColor +"'>Available Updates on " + DateTime.Now.ToShortDateString() + "</font></h1>" +
                "<h2><font color='" + headerColor + "'>Critical Updates: " + cntImp + "</font></h3>" +
                "<h2><font color='" + headerColor + "'>Recommended Updates: " + cntRec + "</font></h2>" +
                "</header>" +
                "<table>" +
                    "<thead>" +
                        "<tr>" +
                            "<th bgcolor =  '" + tableHeaderColor + "' align='center'><font color='" + tableHeaderFontColor  + "' > Server </font></th>" +
                            "<th bgcolor = '" + tableHeaderColor + "' align='center'><font color='" + tableHeaderFontColor + "' > Last Install Date </font ></th>" +
                            "<th bgcolor = '" + tableHeaderColor + "' align='center'><font color='" + tableHeaderFontColor + "' > Important Updates </font></th>" +
                            "<th bgcolor = '" + tableHeaderColor + "' align='center'><font color='" + tableHeaderFontColor + "' > Recommended Updates </font></th>" +
                            "<th bgcolor = '" + tableHeaderColor + "' align='center'><font color='" + tableHeaderFontColor + "' > Error Message </font></th>" +
                        "</tr>" +
                    "</thead>" +
                    "<tbody>";

            // Build table from output servers list.
            foreach (Server s in servers)
            {
                string lastUpdate = "";
                string numImpUpdates = "";
                string numRecUpdates = "";

                //Check if it is an error and set the update count to blank.
                if (String.IsNullOrEmpty(s.Error))
                {
                    lastUpdate = s.LastUpdate.ToShortDateString();
                    numImpUpdates = s.NumImpUpdates.ToString();
                    numRecUpdates = s.NumRecUpdates.ToString();
                }

                //Build row with server info
                textBody += "<tr>" +
                                 "<td>" + s.ServerName + "</td>" +
                                 "<td align='center'>" + lastUpdate + "</td>" +
                                 "<td align='right'>" + numImpUpdates + "</td>" +
                                 "<td align='right'>" + numRecUpdates + "</td>" +
                                 "<td id='errorStyle'>" + s.Error + "</td>" +
                             "</tr>";
            }

            //closing tags
            textBody += "</tbody></table></html>";

            return textBody;
        }

        /*
         * Email results to defined users via smtp.
         */
        public static void SendMailwithTable(String subject, String fromAddress, String toAddress, List<Server> servers)
        {
            try
            {
                //Client connection info
                SmtpClient client = new SmtpClient();
                client.Host = emailHost;
                client.Port = emailPort;

                //create Message
                MailAddress from = new MailAddress(fromAddress);
                MailAddress to = new MailAddress(toAddress);
                MailMessage message = new MailMessage(from, to);

                /* //If using Gmail and port 587 configure the below.  Should only be used for testing.
                 * //Less secure apps needs be turned on here: https://myaccount.google.com/lesssecureapps for the account.  
                 * //Turn off when testing complete.
                client.UseDefaultCredentials = false;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(fromAddress, "your_password");
                */

                //Build Body and Subject enabling HTML
                message.Body = FormatTableForMailSubject(servers);
                message.Body += Environment.NewLine;
                message.IsBodyHtml = true;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Subject = subject;
                message.SubjectEncoding = System.Text.Encoding.UTF8;

                //Send message and dispose.
                client.Send(message);
                message.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Sending Message: " + ex.Message);
                WriteErrorLog(ex.Message);
            }
        }

        /*
         * Method ensures that the Error\Debug directory 
         * exists and if not attempts to create it. Then
         * writes the error for later review.
         */
        private static void WriteErrorLog(string error)
        {
            if (!Directory.Exists(debugDir))
            {
                Directory.CreateDirectory(debugDir);
            }

            if (Directory.Exists(debugDir))
            {
                string filePath = debugDir + "ServerUpdateCheckLog.txt";
                File.AppendAllText(filePath, error);
            }
        }

        /*
         * Decomissioned due to RegKey being inconsistient in my 
         * environment.  Requires remote registry service to be
         * running on target machine and the account running the
         * program to have appropriate access.
         */
        private static DateTime GetLastInstall(string serverName)
        {
            try
            {
                RegistryKey environmentKey = RegistryKey.OpenRemoteBaseKey(
                   RegistryHive.LocalMachine, serverName).OpenSubKey(
                   "Software");
                RegistryKey key = environmentKey.OpenSubKey(@"Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\Results\Install");
                if (key != null)
                {
                    Object o = key.GetValue("LastSuccessTime");
                    if (o != null)
                    {
                        DateTime lastInstallDate;
                        DateTime.TryParse(o.ToString(), out lastInstallDate);

                        if (lastInstallDate != DateTime.MinValue)
                        {
                            return lastInstallDate;
                        }

                    }
                }
            }

            catch (Exception ex)
            {
                output += ex.Message + System.Environment.NewLine;
            }


            return DateTime.MinValue;
        }

        /*
         * Returns server name and last update install date.  Decomissoned
         * due to RegKey being inconsistent.         
         */
        private static string GetBasicServerOutput()
        {
            using (DirectoryEntry de = new
                DirectoryEntry("LDAP://" + serverOU))
            {
                DirectorySearcher src = new DirectorySearcher();
                src.SearchRoot = de;
                src.SearchScope = System.DirectoryServices.SearchScope.Subtree;
                src.Filter = "(objectCategory=computer)";
                List<Server> list = new List<Server>();
                //ArrayList list = new ArrayList();
                SearchResultCollection res = src.FindAll();
                foreach (SearchResult sc in res)
                {
                    foreach (string myCollection in sc.Properties["cn"])
                    {
                        list.Add(new Server(myCollection, GetLastInstall(myCollection)));

                        if (debug)
                            File.AppendAllText(debugFilePath, myCollection);
                    }
                }

                var myList = list.ToList();
                myList.Sort(delegate (Server c1, Server c2) { return c1.LastUpdate.CompareTo(c2.LastUpdate); });
                string output = "";
                foreach (Server s in myList)
                    output += s.ServerName + "," + s.LastUpdate + System.Environment.NewLine;

                return output;
            }
        }

        /*
         * RegKey Email method.  Decommisioned.
         */
        public static void SendMail(String body, String subject, String fromAddress, String toAddress)
        {
            try
            {
                SmtpClient client = new SmtpClient();
                client.Host = emailHost;
                MailAddress from = new MailAddress(fromAddress);
                // Set destinations for the e-mail message.
                MailAddress to = new MailAddress(toAddress);
                // Specify the message content.
                MailMessage message = new MailMessage(from, to);
                message.Body = body;
                message.Body += Environment.NewLine;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Subject = subject;
                message.SubjectEncoding = System.Text.Encoding.UTF8;

                client.Send(message);
                // Clean up.
                message.Dispose();
            }
            catch
            {

            }
        }

    }
}
