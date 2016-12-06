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
    public partial class StationSetting : Form
    {
        public StationSetting()
        {
            InitializeComponent();
        }

        private void StationSetting_Load(object sender, EventArgs e)
        {
            string strsql = "exec sp_StationList";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(strsql);
            StationSetCombo.DataSource = ds.Tables[0];
            StationSetCombo.DisplayMember = "Station";
        }

        private void ModelSave_Click(object sender, EventArgs e)
        {
            string StrSql = $"sp_UpdateStation '{StationSetCombo.Text }','{Dns.GetHostName()}'";
            if (AdoInterface.InsertData(StrSql)==1)
            {
                MessageBox.Show("Station绑定成功");
                Close();
                
            }
            
        //    Properties.Settings.Default.Station = StationSetCombo.Text;
        //    Properties.Settings.Default.Save();
        //    Close();
        }
    }
}
