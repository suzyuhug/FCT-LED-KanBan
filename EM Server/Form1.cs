using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace EM_Server
{
    public partial class Form1 : Form
    {
        SocketManager _sm = null;
        string ip = "127.0.0.1";
        int port = 1113;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _sm = new SocketManager(ip, port);
            _sm.OnReceiveMsg += OnReceiveMsg;
            _sm.OnConnected += OnConnected;
            _sm.OnDisConnected += OnDisConnected;
            _sm.Start();
            init();
        }

        void init()
        {
          
            sviptext.Text = ip;
            svport.Text = port.ToString();
            svstatus.Text = "正常启动";
        }

        public void OnReceiveMsg(string ip)
        {
            byte[] buffer = _sm._listSocketInfo[ip].msgBuffer;
            string msg = Encoding.ASCII.GetString(buffer).Replace("\0", "");
            if (txtMsg.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    txtMsg.Text += AppendReceiveMsg(msg, ip);
                    MessageBox.Show(AppendReceiveMsg(msg, ip));
                }));
            }
            else
            {
                txtMsg.Text += AppendReceiveMsg(msg, ip);
            }
        }
        public void OnConnected(string clientIP)
        {
            string ipstr = clientIP.Split(':')[0];
            string portstr = clientIP.Split(':')[1];
            if (txtMsg.InvokeRequired)
            {
                this.Invoke(new Action(() => {
                    txtMsg.Text += clientIP + "已连接至本机\r\n";
                    object obj = new { Value = clientIP, Text = clientIP };
                    //cbClient.Items.Add(obj);
                    //cbClient.DisplayMember = "Value";
                    //cbClient.ValueMember = "Text";
                    //cbClient.SelectedItem = obj;
                }));
            }
            else
            {
                txtMsg.Text += clientIP + "已连接至本机\r\n";
            }
        }
        public void OnDisConnected(string clientIp)
        {
            if (txtMsg.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    txtMsg.Text += clientIp + "已经断开连接！\r\n";
                    object obj = new { Value = clientIp, Text = clientIp };
                    //cbClient.Items.Remove(obj);
                }));
            }
            else
            {
                txtMsg.Text += clientIp + "已经断开连接！\r\n";
            }
        }
        public string AppendReceiveMsg(string msg, string ipClient)
        {
            return GetDateNow() + "  " + "[接收" + ipClient + "]  " + msg + "\r\n";

        }
        public string GetDateNow()
        {
            return DateTime.Now.ToString("HH:mm:ss");
        }
    }
}
