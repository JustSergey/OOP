using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace OOPServer
{
    class Program
    {
        const string default_ip = "127.0.0.1";
        const int default_port = 3344;
        const string port_info_path = "port.ini";
        const string result_file_name = "result.dat";    
        public enum MessageType { GetResult, SendFile };

        static Dictionary<MessageType, Action<NetworkStream>> RequestHandlers
            = new Dictionary<MessageType, Action<NetworkStream>>
            { {MessageType.GetResult,CreateAndSendResultFile},{ MessageType.SendFile,AcceptAndSaveFile } };

        static void Main(string[] args)
        {
            int port = ParsePort(port_info_path);
            TcpListener server = new TcpListener(IPAddress.Any, port);
            server.Start();

            while (true)
            {
                HandleRequest(server);                
            }
        }

        static void HandleRequest(TcpListener server)
        {
            TcpClient client = server.AcceptTcpClient();
            using (NetworkStream stream = client.GetStream())
            {
                MessageType messageType = (MessageType)BitConverter.ToInt32(ParseData(stream), 0);
                if (RequestHandlers.Keys.Contains(messageType))
                {
                    RequestHandlers[messageType](stream);
                }
                else
                {
                    Console.WriteLine("Error Request");
                }
                
            }
            client.Close();
        }

        static byte[] ParseData(NetworkStream stream)
        {
            byte[] buffer_for_int = new byte[4];
            stream.Read(buffer_for_int, 0, 4);
            return buffer_for_int;            
        }

        static void AcceptAndSaveFile(NetworkStream stream)
        {
            int quarter_index = BitConverter.ToInt32(ParseData(stream), 0);
            int branch_index = BitConverter.ToInt32(ParseData(stream), 0);
            List<byte> data = new List<byte>();
            do
            {
                byte[] buffer = new byte[256];
                int bytes_count = stream.Read(buffer, 0, buffer.Length);
                for (int i = 0; i < bytes_count; i++)
                    data.Add(buffer[i]);
            } while (stream.DataAvailable);

            string file_name = "br" + branch_index + "q" + quarter_index + ".dat";
            File.WriteAllBytes(file_name, data.ToArray());
            Console.WriteLine("File " + file_name + " was received");
        }

        static void CreateAndSendResultFile(NetworkStream stream)
        {
            int quarter_index = BitConverter.ToInt32(ParseData(stream), 0);//пока год = 5
            // пока для тестов
            if (quarter_index == 5)
            {

                //string[][] files = new string[4][];
                //for( var i = 0;i<4;i++)
                //files[i] = Directory.GetFiles(Directory.GetCurrentDirectory(), "br?q" + i + ".dat");
                
                //Main_Form main_form = new Main_Form(files);
                //main_form.ShowDialog();
                //byte[] file = File.ReadAllBytes(result_file_name);
                //stream.Write(file, 0, file.Length);
                //Thread.Sleep(200);
                //Console.WriteLine("File result.dat was sent");
                //File.Delete(result_file_name);
            }

            else
            {
                string[] files = Directory.GetFiles(Directory.GetCurrentDirectory(), "br?q" + quarter_index + ".dat");
                if (files.Length > 0)
                {
                    Main_Form main_form = new Main_Form(files);
                    main_form.ShowDialog();
                    byte[] file = File.ReadAllBytes(result_file_name);
                    stream.Write(file, 0, file.Length);
                    Thread.Sleep(200);
                    Console.WriteLine("File result.dat was sent");
                    File.Delete(result_file_name);
                }
            }
        }

        static int ParsePort(string port_info_path)
        {
            if (!File.Exists(port_info_path))
            {
                File.WriteAllText(port_info_path, default_port.ToString());
                return default_port;
            }
            if (!int.TryParse(File.ReadAllText(port_info_path), out int port))
                return -1;
            return port;
        }
    }
}
