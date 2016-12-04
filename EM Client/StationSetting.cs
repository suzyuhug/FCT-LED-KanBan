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
            Properties.Settings.Default.Station = StationSetCombo.Text;
            Properties.Settings.Default.Save();
            Close();
        }
    }
}
