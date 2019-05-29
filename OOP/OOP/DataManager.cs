using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.IO;

namespace OOP
{
    public static class DataManager
    {
        public enum MessageType { GetResult, SendFile };

        private static TcpClient client;
        public static string BranchName;
        public static int BranchIndex = -1;
        public static int QuarterIndex = -1;

        public const string ip_info_path = "ip.ini";
        public const string branches_info_path = "branches.inf";

        public static bool ConnectToServer(IPEndPoint address)
        {
            client = new TcpClient();
            return NetManager.ConnectToServer(client, address);
        }

        public static void SendRequest(MessageType messageType, string file_path)
        {
            NetworkStream stream = client.GetStream();

            byte[] type = BitConverter.GetBytes((int)messageType);
            stream.Write(type, 0, type.Length);

            byte[] index = BitConverter.GetBytes(QuarterIndex);
            stream.Write(index, 0, index.Length);

            if (messageType == MessageType.SendFile)
            {
                index = BitConverter.GetBytes(BranchIndex);
                stream.Write(index, 0, index.Length);

                byte[] file = File.ReadAllBytes(file_path);
                stream.Write(file, 0, file.Length);
            }
        }

        public static void GetResponse(string result_file_path)
        {
            NetworkStream stream = client.GetStream();
            List<byte> data = new List<byte>();
            do
            {
                byte[] buffer = new byte[256];
                int bytes_count = stream.Read(buffer, 0, buffer.Length);
                for (int i = 0; i < bytes_count; i++)
                    data.Add(buffer[i]);
            } while (stream.DataAvailable);

            File.WriteAllBytes(result_file_path, data.ToArray());
        }

        public static void Serialize(DataGridView[] Tables, string file_path)
        {
            using (FileStream file = File.Create(file_path))
            {
                for (int i = 0; i < Tables.Length; i++)
                {
                    byte[] buffer = BitConverter.GetBytes(Tables[i].ColumnCount);
                    file.Write(buffer, 0, buffer.Length);
                    buffer = BitConverter.GetBytes(Tables[i].RowCount);
                    file.Write(buffer, 0, buffer.Length);
                    for (int y = 0; y < Tables[i].RowCount; y++)
                    {
                        for (int x = 0; x < Tables[i].ColumnCount; x++)
                        {
                            buffer = Encoding.Unicode.GetBytes(Tables[i][x, y].Value.ToString());
                            file.Write(buffer, 0, buffer.Length);
                            file.WriteByte(0x02);
                            file.WriteByte(0xA8);
                        }
                    }
                }
            }
        }

        public static void Deserialize(DataGridView[] Tables, string file_path)
        {
            if (!File.Exists(file_path))
                return;

            using (FileStream file = File.OpenRead(file_path))
            {
                for (int i = 0; i < Tables.Length; i++)
                {
                    file.Position += 8;
                    for (int y = 0; y < Tables[i].RowCount; y++)
                    {
                        for (int x = 0; x < Tables[i].ColumnCount; x++)
                        {
                            List<byte> buffer = new List<byte>(byte.MaxValue);
                            byte[] bytes = new byte[2];
                            while (file.Read(bytes, 0, 2) == 2)
                            {
                                if (bytes[0] == 0x02 && bytes[1] == 0xA8)
                                    break;
                                buffer.AddRange(bytes);
                            }
                            Tables[i][x, y].Value = Encoding.Unicode.GetString(buffer.ToArray());
                        }
                    }
                }
            }
        }
    }
}
