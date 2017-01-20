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
    public partial class FaFrm : Form
    {
        public FaFrm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AdoInterface.FrmfailMes = comboBox1.Text;
            Close();
        }

        private void FaFrm_Load(object sender, EventArgs e)
        {
            string Strsql = "exec sp_FailList";
            DataSet ds = new DataSet();
            ds = AdoInterface.GetDataSet(Strsql);
            comboBox1.DataSource = ds.Tables[0];
            comboBox1.DisplayMember = "FailMessage";
            AdoInterface.FrmfailMes ="";
        }
    }
}
