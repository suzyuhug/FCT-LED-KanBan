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

      

        private void Form1_Load(object sender, EventArgs e)
        {
      
            Station();
            InitSocket();
            modellist();
            pageset();
            
           

        }
        private void Station()//查询当前电脑绑定的Station
        {
            string Strsql = $"exec sp_Station_Find '{Dns.GetHostName()}'";
            StationLab.Text = AdoInterface.Readstr(Strsql);

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
            string Strsql = $"exec sp_ModelList '{StationLab.Text }'";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(Strsql);
            Modelcombo.DataSource = ds.Tables[0];
            Modelcombo.DisplayMember = "Model";

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
            if (!_scm._socket.Connected) return;
            _scm._isConnected = false;
            _scm.SendMsg("\0\0\0faild");
            _scm._socket.Shutdown(System.Net.Sockets.SocketShutdown.Both);
            _scm._socket.Close();
        }

        private void 站别绑定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StationSetting SetFrm = new EM_Client.StationSetting();
            SetFrm.ShowDialog();
            //StationLab.Text = Properties.Settings.Default.Station;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            panel2.Width = 500;
            panel2.Height = 430;
            panel2.Top = 60;
            panel2.Left = 15;
            panel1.Visible = false;
            panel2.Visible = true;


            string StrSql = $"exec sp_AssyStep '{Modelcombo.Text}','{StationLab.Text}'";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(StrSql);
            GridView.DataSource = ds.Tables[0];


            rockontime();



        }
        private void rockontime()//进度条计时
        {
            string StrSql = $"exec sp_QueryH '{Modelcombo.Text}'";
            int t = int.Parse (AdoInterface.Readstr(StrSql));
            label4.Text = $"组装时长：{t/60} 小时";
            PB.Maximum = t;
            PB.Value = t;
            timer1.Interval = 2000;
            timer1.Start();


        }
        string tempgvid = null;
        private void GridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (GridView.RowCount > 0)
            {
                if (e.ColumnIndex == 3)
                {
                    if (label3.Text=="")
                    {

                        if (GridView.CurrentRow.Cells["EntBut"].Value.ToString() != "完成组装")
                        {
                            for (int i = 0; i < GridView.RowCount; i++)
                            {
                                if (GridView.Rows[i].Cells["EntBut"].Value.ToString() == "正在组装")
                                {
                                    GridView.Rows[i].Cells["EntBut"].Value = "完成组装";
                                    GridView.Rows[i].Cells["bs"].Value = imageList1.Images[1];
                                }

                            }



                            GridView.CurrentRow.Cells["EntBut"].Value = "正在组装";
                            GridView.CurrentRow.Cells["bs"].Value = imageList1.Images[0];
                            tempgvid = GridView.CurrentRow.Cells["ID"].Value.ToString();
                            _scm.SendMsg($"Operational#{StationLab.Text}#{GridView.CurrentRow.Cells["ID"].Value.ToString()}#100%");

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


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            FaFrm Fa = new EM_Client.FaFrm();
            Fa.ShowDialog();
            string str = AdoInterface.FrmfailMes;
            label3.Text = str;
            _scm.SendMsg($"Unusual#{StationLab.Text}#{Perlabel.Text}#{str}");

        }
        int num;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (PB.Value > 0)
            {
                PB.Value = PB.Value - 1;
                //  Perlabel.Text = ((PB.Value / PB.Maximum) * 100).ToString();
                double percent = (double)PB.Value / PB.Maximum;
                Perlabel.Text =percent.ToString ("0.0%");
                num++;
                if (num==10)
                {
                    num = 0;
                    _scm.SendMsg($"Operational#{StationLab.Text}#{tempgvid}#{Perlabel.Text}");
                }
             
               

                
            }
            else
            {
                timer1.Enabled = false;
                MessageBox.Show("down");
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            panel2.Visible = false;
            pageset();
            PB.Maximum = 100;PB.Value = 100;
            Perlabel.Text = "100%";
            _scm.SendMsg($"Completed#{StationLab.Text}#{tempgvid}#{Perlabel.Text}");
        }
    }
}
