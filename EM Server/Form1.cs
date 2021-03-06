﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using Demo;

namespace EM_Server
{
    public partial class Form1 : Form
    {
        SocketManager _sm = null;

       string ip ="10.194.48.150";
        //string ip = "10.194.40.65";
       // string ip = "127.0.0.1";


        int port = 1113;

     

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            this.Text = this.Text + "-" + Application.ProductVersion;
            _sm = new SocketManager(ip, port);
            _sm.OnReceiveMsg += OnReceiveMsg;
            _sm.OnConnected += OnConnected;
            _sm.OnDisConnected += OnDisConnected;
            _sm.Start();
            init();
        }
        private void AssyLoad(string St, int id, string tjd, string mes,string strcolor)//发送组装信息
        {
            string Strsql = $"exec sp_AssyLoad '{St}','{id}'";
            DataSet ds = new DataSet();
            ds = Adoread.GetDataSet(Strsql);
            if (ds.Tables[0].Rows.Count > 0)
            {

                string tempStation = ds.Tables[0].Rows[0]["Station"].ToString();
                string tempIP = ds.Tables[0].Rows[0]["KanBanIP"].ToString();
                string tempMes;
                if (mes == null)
                {
                    tempMes = ds.Tables[0].Rows[0]["Message"].ToString() + "              ";

                }
                else
                {
                    tempMes = mes;
                }


                string tempjd = tjd;
                listBox1.Items.Add ($"{tempStation}：{tempMes}：{tempjd}");
                ledkanban(tempIP, tempStation, tempMes, tempjd,strcolor);

            }
        }

        void init()
        {
          
            sviptext.Text =$"Server IP:{ip}";
            svport.Text =$"Port:{port.ToString()}";
            svstatus.Text = $"Status：successful";
        }

        public void OnReceiveMsg(string ip)
        {
            byte[] buffer = _sm._listSocketInfo[ip].msgBuffer;
            string msg = Encoding.ASCII.GetString(buffer).Replace("\0", "");
            if (listBox1 .InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                   
                    string str = AppendReceiveMsg(msg, ip);
                    string[] sArray = str.Split('#');
                    if (sArray[0]=="Operational")
                    {
                        AssyLoad(sArray[1], int.Parse(sArray[2]), sArray[3],null,"g");
                    }
                    else if (sArray[0] == "Completed")
                    {
                        AssyLoad(sArray[1], int.Parse(sArray[2]), "100%","Completed                ","g");
                    }
                    else if (sArray[0] == "Unusual")
                    {
                        AssyLoad(sArray[1], 0, sArray [2], sArray[3]+"              ","r");
                       // _scm.SendMsg($"Unusual#{StationLab.Text}#{str}#100%");
                    }
                    else if (sArray[0] == "Waiting")
                    {
                        AssyLoad(sArray[1], 0, sArray[2], sArray[3] + "              ", "g");
                        // _scm.SendMsg($"Unusual#{StationLab.Text}#{str}#100%");
                    }

                }));
            }
            else
            {
                listBox1.Items.Add(AppendReceiveMsg(msg, ip));
            }
        }
        public void OnConnected(string clientIP)
        {
            string ipstr = clientIP.Split(':')[0];
            string portstr = clientIP.Split(':')[1];
            if (listBox1 .InvokeRequired)
            {
                this.Invoke(new Action(() => {
                    listBox1.Items.Add(clientIP + "  已连接至本机");    
                    listView1.Items.Add(ipstr,0);
                   
                }));
            }
            else
            {
                listBox1.Items.Add(clientIP + "  已连接至本机");
            }
        }
        private void deluser(string str)
        {

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].Text == str)
                {
                    listView1.Items.RemoveAt(i);
                }

            }

        }
        public void OnDisConnected(string clientIp)
        {
            if (listBox1 .InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                     listBox1.Items.Add(clientIp + "  已断开连接");
                    string ipstr = clientIp.Split(':')[0];
                    deluser(ipstr);
                }));
            }
            else
            {
                listBox1.Items.Add(clientIp + "  已断开连接");
            }
        }
        public string AppendReceiveMsg(string msg, string ipClient)
        {
            return msg;

        }
        public string GetDateNow()
        {
            return DateTime.Now.ToString("HH:mm:ss");
        }

      

        private void ledkanban(string TempIP,string Tempst,string message,string Jd,string strcolor)
        {
            int nResult;
            LedDll.COMMUNICATIONINFO CommunicationInfo = new LedDll.COMMUNICATIONINFO();
            CommunicationInfo.SendType = 0;
            CommunicationInfo.IpStr = TempIP;
            CommunicationInfo.LedNumber = 1;
            int hProgram;
            hProgram = LedDll.LV_CreateProgram(64, 32, 2);
            nResult = LedDll.LV_AddProgram(hProgram, 1, 0, 10);
            if (nResult != 0)
            {
                string ErrStr;
                ErrStr = LedDll.LS_GetError(nResult);
                MessageBox.Show(ErrStr);
                return;
            }
            LedDll.AREARECT AreaRect = new LedDll.AREARECT();
            AreaRect.left = 0;
            AreaRect.top = 0;
            AreaRect.width = 64;
            AreaRect.height = 16;
            LedDll.FONTPROP FontProp = new LedDll.FONTPROP();
            FontProp.FontName = "Arial";
            FontProp.FontSize = 10;
            if (strcolor =="r")
            {
FontProp.FontColor = LedDll.COLOR_RED;
            }
            else if(strcolor=="g")
            {
                FontProp.FontColor = LedDll.COLOR_GREEN;
            }
            
            FontProp.FontBold = 0;
            nResult = LedDll.LV_QuickAddSingleLineTextArea(hProgram, 1, 1, ref AreaRect, LedDll.ADDTYPE_STRING, message , ref FontProp, 4);
            AreaRect.left = 0;
            AreaRect.top = 16;
            AreaRect.width = 64;
            AreaRect.height = 16;
            FontProp.FontSize = 10;
            LedDll.LV_AddSingleLineTextToImageTextArea(hProgram, 1, 2, ref AreaRect, 0);
            nResult = LedDll.LV_AddStaticTextToImageTextArea(hProgram, 1, 2, LedDll.ADDTYPE_STRING, Tempst , ref FontProp, 1, 2, 1);
            nResult = LedDll.LV_AddProgram(hProgram, 2, 0, 5);
            AreaRect.left = 0;
            AreaRect.top = 16;
            AreaRect.width = 64;
            AreaRect.height = 16;
            FontProp.FontName = "Arial";
            FontProp.FontSize = 10;
            FontProp.FontColor = LedDll.COLOR_GREEN;
            FontProp.FontBold = 0;
            LedDll.LV_AddSingleLineTextToImageTextArea(hProgram, 2, 2, ref AreaRect, 0);
            nResult = LedDll.LV_AddStaticTextToImageTextArea(hProgram, 2, 2, LedDll.ADDTYPE_STRING, Jd, ref FontProp, 1, 2, 1);


            // nResult = LedDll.LV_QuickAddSingleLineTextArea(hProgram, 2, 2, ref AreaRect, LedDll.ADDTYPE_STRING, Jd , ref FontProp, 4);
            AreaRect.left = 0;
            AreaRect.top = 0;
            AreaRect.width = 64;
            AreaRect.height = 16;
            FontProp.FontSize = 10;
            LedDll.LV_AddSingleLineTextToImageTextArea(hProgram, 2, 1, ref AreaRect, 0);
            nResult = LedDll.LV_AddStaticTextToImageTextArea(hProgram, 2, 1, LedDll.ADDTYPE_STRING, "Completed" , ref FontProp, 1, 2, 1);
            nResult = LedDll.LV_Send(ref CommunicationInfo, hProgram);
            LedDll.LV_DeleteProgram(hProgram);
            if (nResult != 0)
            {
                string ErrStr;
                ErrStr = LedDll.LS_GetError(nResult);
             
                listBox1.Items.Add(ErrStr );
            }
            else
            {
                listBox1.Items.Add(TempIP +"    LED发发送成功");
            }

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason==CloseReason.UserClosing )
            {
                e.Cancel = true;
                this.Hide();
            }
        }

        private void 显示窗体ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Show();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
            Application.Exit();
        }
    }
}
