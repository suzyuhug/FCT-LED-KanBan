using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace EM_Client
{
   


    class AdoInterface
    {
        static  string SqlData = "server=10.194.48.150\\MySQL;database=FCT_LED_KanBan;uid=sa;pwd=Aa123456";
        public static string FrmfailMes = null;

        public static DataSet  GetDataSet(string strSql)
        {          
                SqlConnection cn = new SqlConnection(SqlData);
                cn.Open();          
                SqlCommand cmd = new SqlCommand(strSql , cn);
                SqlDataAdapter dp = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                dp.Fill(ds);                                
                cn.Close();
            cn.Dispose();
                return ds;         
        }

        public static string Readstr(string strSql)
        {
            string str;
            SqlConnection cn = new SqlConnection(SqlData);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strSql, cn);
            SqlDataAdapter dp = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dp.Fill(ds);
            cn.Close();
            cn.Dispose();
            if (ds.Tables[0].Rows.Count >0)
            {
                str = ds.Tables[0].Rows[0][0].ToString();
                return str;
            }
            return null;

        }
        public static int InsertData(string strSql)
        {
            SqlConnection cn = new SqlConnection(SqlData);        
            SqlCommand cmd = new SqlCommand(strSql, cn);
            cn.Open();
            cmd.ExecuteNonQuery();
            cn.Close();
            cn.Dispose();
            return 1;
        }

    }
}
