using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;

namespace OOP
{
    public static class NetManager
    {
        const string default_ip = "127.0.0.1";
        const int default_port = 3344;

        public static bool ConnectToServer(TcpClient client, IPEndPoint address)
        {
           return client.ConnectAsync(address.Address, address.Port).Wait(2000);
        }

        public static IPEndPoint GetAddress(string file_path)
        {
            IPEndPoint address = new IPEndPoint(IPAddress.Parse(default_ip), default_port);

            if (!File.Exists(file_path))
                CreatIpConfigFile(file_path, address);

            string ipPort = File.ReadAllText(file_path);
            return IsCorrectAddress(ipPort) ? ParseIp(ipPort) : address;
        }

        public static void CreatIpConfigFile(string file_path, IPEndPoint address)
        {
            File.WriteAllText(file_path, address.Address.ToString() + ":" + address.Port.ToString());
        }

        public static IPEndPoint ParseIp(string address)
        {
            string[] ipAndPort = address.Split(':');
            return new IPEndPoint(IPAddress.Parse(ipAndPort[0]), int.Parse(ipAndPort[1]));
        }

        public static bool IsCorrectAddress(string address)
        {
            string[] ipAndPort = address.Split(':');
            return ipAndPort.Length == 2 && IsCorrectIp(ipAndPort[0]) && IsCorrectPort(ipAndPort[1]);
        }

        public static bool IsCorrectIp(string ipAndPort)
        {
            IPAddress ip;
            return IPAddress.TryParse(ipAndPort, out ip);
        }

        public static bool IsCorrectPort(string ipAndPort)
        {
            int port;
            return int.TryParse(ipAndPort, NumberStyles.None, NumberFormatInfo.CurrentInfo, out port);
        }
    }
}
