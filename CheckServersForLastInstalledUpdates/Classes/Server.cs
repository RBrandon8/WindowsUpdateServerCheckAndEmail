using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CheckServersForLastInstalledUpdates.Classes
{
    class Server
    {
        public string ServerName { get; set; }
        public DateTime LastUpdate { get; set; }
        public int NumImpUpdates { get; set; }
        public int NumRecUpdates { get; set; }
        public string Error { get; set; }

        public Server()
        {
            Error = "";
        }

        public Server(String serverName, DateTime lastUpdate)
        {
            ServerName = serverName;
            LastUpdate = lastUpdate;
            Error = "";
        }

        public Server(String serverName, DateTime lastUpdate, int numImpUpdates, int numRecUpdates)
        {
            ServerName = serverName;
            LastUpdate = lastUpdate;
            NumImpUpdates = numImpUpdates;
            NumRecUpdates = numRecUpdates;
            Error = "";
        }



        public Server(String serverName, string error)
        {
            ServerName = serverName;
            LastUpdate = DateTime.MinValue;
            NumImpUpdates = 0;
            NumRecUpdates = 0;
            Error = error;
        }
    }
}
