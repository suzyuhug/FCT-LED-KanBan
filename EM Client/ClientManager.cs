using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace EM_Client
{
    class ClientManager
    {
        public Socket _socket = null;
        public EndPoint endPoint = null;
        public SocketInfo socketInfo = null;
        public bool _isConnected = false;

        public delegate void OnConnectedHandler();
        public event OnConnectedHandler OnConnected;
        public event OnConnectedHandler OnFaildConnect;
        public delegate void OnReceiveMsgHandler();
      

        public ClientManager(string ip, int port)
        {
            IPAddress _ip = IPAddress.Parse(ip);
            endPoint = new IPEndPoint(_ip, port);
            _socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
        }

        public void Start()
        {
            _socket.BeginConnect(endPoint, ConnectedCallback, _socket);
            _isConnected = true;
          
        }

       

        public void ConnectedCallback(IAsyncResult ar)
        {
            Socket socket = ar.AsyncState as Socket;
            if (socket.Connected)
            {
                if (this.OnConnected != null) OnConnected();
            }
            else
            {
                if (this.OnFaildConnect != null) OnFaildConnect();
            }
        }

        public void SendMsg(string msg)
        {
            byte[] buffer = Encoding.ASCII.GetBytes(msg);
            _socket.Send(buffer);
        }

        public class SocketInfo
        {
            public Socket socket = null;
            public byte[] buffer = null;

            public SocketInfo()
            {
                buffer = new byte[1024];
            }
        }


    }
}
