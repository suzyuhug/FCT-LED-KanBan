using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace EM_Client
{
    public partial class Form1 : Form
    {
        ClientManager _scm = null;
        string ip = "10.194.48.150";
        int port = 1113;
        public Form1()
        {
            InitializeComponent();
        }
        bool cl = false;
        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateClass.UpdateFrom("FCT-LED-Client");
            this.Text = this.Text + "-" + Application.ProductVersion;
            Station();
            InitSocket();
            modellist();
            string Strsql = $"exec sp_Busy '{StationLab.Text}'";
            string statusID = AdoInterface.Readstr(Strsql);
            if (statusID == "0")
            {
                pageset();
            }
            else if (statusID == "1")
            {
                loadtempmodel();
                button2_Click(sender, e);
                loadtime();
            }
            cl = true;
        }
        private void loadtempmodel()//加载临时Model
        {
            string StrSql = $"exec sp_TempModel '{StationLab .Text}'";
            string str = AdoInterface.Readstr(StrSql);
            if (str != null)
            {
                 label2.Text =str;                         
            }
        }
        private void loadtime()//加载剩余时间
        {
            string StrSql = $"exec sp_LoadTime '{StationLab.Text}'";
            string str = AdoInterface.Readstr(StrSql);
            if (str != null)
            {
                int t = int.Parse(str);               
                PB.Value = PB.Value-t;
            }
        }
        private void Tryconnect()//尝试再次连接
        {
            panel3.Width = 317;
            panel3.Height = 136;
            panel3.Top = 170;
            panel3.Left = 250;
            timer2.Interval = 100;
            timer2.Start();
            panel3.Visible = true;
        }
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (PBCS.Value < 100)
            {
                PBCS.Value += 1;
            }
            else
            {
                InitSocket();
                PBCS.Value = 0;
                if (svstatus.Text == "Status：successful")
                {
                    panel3.Visible = false;
                    timer2.Enabled = false;
                }
            }
        }
        private void Station()//查询当前电脑绑定的Station
        {
            string Strsql = $"exec sp_Station_Find '{Dns.GetHostName()}'";
            string str = AdoInterface.Readstr(Strsql);
            if (str != null)
            {
                StationLab.Text = str;
            }
        }

        private void pageset()//页面布局设置
        {
            panel1.Width = 365;
            panel1.Height = 165;
            panel1.Top = 80;
            panel1.Left = 30;
            panel1.Visible = true;
        }
        private void modellist()//加载Model comlist数据
        {
            string Strsql = $"exec sp_ModelList '{StationLab.Text}'";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(Strsql);
            if (ds != null)
            {
                Modelcombo.DataSource = ds.Tables[0];
                Modelcombo.DisplayMember = "Model";                              
            }
        }
        private void sendmessage(string message)
        {
            if (svstatus.Text == "Status：successful")
            {
                try
                {
                    _scm.SendMsg(message);
                }
                catch (Exception)
                {
                    MessageBox.Show("服务器断开连接，请重试！");
                    svstatus.Text = "连接失败";
                    Tryconnect();
                }
            }
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
                this.Invoke(new Action(() =>
                {
                    serstatus.Text = "服务器连接成功";
                    svip.Text = $"Server IP：{ip}";
                    svport.Text = $"Server Port：{ port.ToString()}";
                    svstatus.Text = "Status：successful";
                }));
            }
            else
            {
                serstatus.Text = "服务器连接成功";
            }
        }
        public void OnFaildConnect()//连接失败
        {
            if (serstatus.InvokeRequired)
            {
                this.Invoke(new Action(() =>
                {
                    serstatus.Text = "服务器连接失败";
                    svstatus.Text = "连接失败";
                }));
            }
            else
            {
                serstatus.Text = "服务器连接失败";
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (cl)
            {
                if (!_scm._socket.Connected) return;
                _scm._isConnected = false;
                _scm.SendMsg("\0\0\0faild");
                _scm._socket.Shutdown(System.Net.Sockets.SocketShutdown.Both);
                _scm._socket.Close();

            }
           
        }

        private void 站别绑定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StationSetting SetFrm = new EM_Client.StationSetting();
            SetFrm.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel2.Width = 500;
            panel2.Height = 430;
            panel2.Top = 60;
            panel2.Left = 15;
            panel1.Visible = false;
            panel2.Visible = true;
            if (label2.Text == "Station")
            {
                label2.Text = Modelcombo.Text;
            }          
            string StrSql = $"exec sp_AssyStep '{label2.Text}','{StationLab.Text}'";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(StrSql);
            GridView.DataSource = ds.Tables[0];
            for (int i = 0; i < GridView.Rows.Count ; i++)
            {
                if (GridView.Rows[i].Cells["EntBut"].Value.ToString() == "正在组装")
                {
                    tempgvid = GridView.Rows[i].Cells["ID"].Value.ToString();
                }
                switch (GridView.Rows[i].Cells["EntBut"].Value.ToString())
                {
                    case "正在组装":
                        GridView.Rows[i].Cells["bs"].Value = imageList1.Images[0];
                       
                        break;
                    case "完成组装":
                        GridView.Rows[i].Cells["bs"].Value = imageList1.Images[1];
                        break;
                }
            }
            rockontime();
        }
        private void rockontime()//进度条计时
        {
            string StrSql = $"exec sp_QueryH '{label2.Text}'";
            string str = AdoInterface.Readstr(StrSql);
            if (str != null)
            {
                int t = int.Parse(str);              
                TimeSpan ts = new TimeSpan(0, t, 0);
                label7.Text = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
              //  label5.Text = label7.Text;
               // Perlabel.Text = "0%";
                PB.Maximum = t;
                PB.Value = t;
                num = 8;
                time1click();
                timer1.Interval = 60000;
                timer1.Start();
            }
        }
        string tempgvid = null;
        private void GridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (GridView.RowCount > 0)
            {
                if (e.ColumnIndex == 3)
                {
                    if (svstatus.Text == "Status：successful")
                    {
                        if (label3.Text == "")
                        {
                            string StrSql = null;
                            if (GridView.CurrentRow.Cells["EntBut"].Value.ToString() != "完成组装")
                            {
                                for (int i = 0; i < GridView.RowCount; i++)
                                {
                                    if (GridView.Rows[i].Cells["EntBut"].Value.ToString() == "正在组装")
                                    {
                                        GridView.Rows[i].Cells["EntBut"].Value = "完成组装";
                                        GridView.Rows[i].Cells["bs"].Value = imageList1.Images[1];
                                        StrSql = $"exec sp_UpdateTempStepStatus '{int.Parse(GridView.Rows[i].Cells["ID"].Value.ToString())}','完成组装'";
                                        if (AdoInterface.InsertData(StrSql) == 0)
                                        {
                                            MessageBox.Show("数据库连接失败，无法更新状态！");
                                        }
                                    }
                                }

                                GridView.CurrentRow.Cells["EntBut"].Value = "正在组装";
                                GridView.CurrentRow.Cells["bs"].Value = imageList1.Images[0];
                                tempgvid = GridView.CurrentRow.Cells["ID"].Value.ToString();
                                sendmessage($"Operational#{StationLab.Text}#{GridView.CurrentRow.Cells["ID"].Value.ToString()}#{Perlabel.Text}");
                                StrSql = $"exec sp_UpdateTempStepStatus '{int.Parse(tempgvid)}','正在组装'";
                                if (AdoInterface.InsertData(StrSql) == 0)
                                {
                                    MessageBox.Show("数据库连接失败，无法更新状态！");
                                }
                            }
                            else
                            {
                                MessageBox.Show("已经完成不可以再点");
                            }
                        }
                        else
                        {
                            MessageBox.Show("有异常信息在进行中");
                            label3.Text = "";
                            timer1.Enabled = true;
                        }
                    }
                    else
                    {
                        MessageBox.Show("服务端连接失败，请重试！");
                        Tryconnect();
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (svstatus.Text == "Status：successful")
            {
                timer1.Enabled = false;
                FaFrm Fa = new EM_Client.FaFrm();
                Fa.ShowDialog();
                string str = AdoInterface.FrmfailMes;
                label3.Text = str;
                sendmessage($"Unusual#{StationLab.Text}#{Perlabel.Text}#{str}");
                timer3.Interval = 60000;
                timer3.Start();
            }
            else
            {
                MessageBox.Show("服务端连接失败，请重试！");
                Tryconnect();
            }
        }
        int num;
        private void timer1_Tick(object sender, EventArgs e)
        {
            time1click();
        }
        private void time1click()
        {
            if (PB.Value > 0)
            {
                PB.Value = PB.Value - 1;
                double percent = (double)(PB.Maximum - PB.Value) / PB.Maximum;
                Perlabel.Text = percent.ToString("0.0%");
                TimeSpan ts = new TimeSpan(0, PB.Value, 0);
                label5.Text = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
                num++;

                if (num == 10)
                {
                    num = 0;
                    if (svstatus.Text == "Status：successful")
                    {
                        sendmessage($"Operational#{StationLab.Text}#{tempgvid}#{Perlabel.Text}");
                    }
                }
            }
            else
            {
                dome();
                MessageBox.Show("组装完成");

            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            dome();
        }
        private void dome()
        {
            if (svstatus.Text == "Status：successful")
            {
                timer1.Enabled = false;
                panel2.Visible = false;
                pageset();
                PB.Maximum = 100; PB.Value = 100;
                Perlabel.Text = "100%";
                sendmessage($"Completed#{StationLab.Text}#{tempgvid}#{Perlabel.Text}");
                label2.Text = "Station";
                string StrSql = $"exec sp_UpdateStationStatus '{StationLab.Text}'";
                if (AdoInterface.InsertData(StrSql) == 0)
                {
                    MessageBox.Show("数据库连接失败，无法更新状态！");
                }
                label5.Text = "00:00"; label7.Text = "00:00";
            }
            else
            {
                MessageBox.Show("服务端连接失败，请重试！");
                Tryconnect();
            }


        }

        private void 等待组装ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (svstatus.Text == "Status：successful")
            {
                timer1.Enabled = false;                
                sendmessage($"Unusual#{StationLab.Text}#0%#Waiting");
            }
            else
            {
                MessageBox.Show("服务端连接失败，请重试！");
                Tryconnect();
            }
        }
        int yctime = 0;
        private void timer3_Tick(object sender, EventArgs e)
        {
            if (label3.Text != "")
            {
                yctime++;
                button3.Text = $"异 常  {yctime}";
            }
            else
            {
                timer3.Enabled = false;
               
                button3.Text = "异 常";
                string StrSql = $"exec sp_UpdateWaitingTime '{yctime}','{StationLab.Text}'";
                if (AdoInterface.InsertData(StrSql) == 0)
                {
                    MessageBox.Show("数据库连接失败，无法更新状态！");
                }
                yctime = 0;
            }
        }
    }
}
