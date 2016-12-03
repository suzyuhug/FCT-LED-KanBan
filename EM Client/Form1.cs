using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EM_Client
{
    public partial class Form1 : Form
    {
        ClientManager _scm = null;
        string ip = "127.0.0.1";
        int port = 1113;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _scm.SendMsg("5835259");
           
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitSocket();
        }

        private void InitSocket()//用于服务端的连接
        {
            _scm = new ClientManager(ip, port);
            _scm.OnConnected += OnConnected;
            _scm.OnFaildConnect += OnFaildConnect;
            _scm.Start();
        }
        public void OnConnected()//连接成功
        {
            if (serstatus.InvokeRequired)
            {
                this.Invoke(new Action(() => {
                    serstatus.Text = "服务器连接成功";   
                    svip .Text =$"Server IP：{ip}";
                    svport.Text = $"Server Port：{ port.ToString()}";
                    svstatus .Text = "Status：successful";
                }));
            }
            else
            {
                serstatus.Text = "服务器连接成功";
            }
        }
        public void OnFaildConnect()//连接失败
        {
            if (serstatus .InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    serstatus.Text = "服务器连接失败";
                    svstatus .Text = "连接失败";
                }));
            }
            else
            {
                serstatus.Text = "服务器连接失败";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!_scm._socket.Connected) return;
            _scm._isConnected = false;
            _scm.SendMsg("\0\0\0faild");
            _scm._socket.Shutdown(System.Net.Sockets.SocketShutdown.Both);
            _scm._socket.Close();
            
        }
    }
}
