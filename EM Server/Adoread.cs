﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace EM_Server
{
    class Adoread
    {

        static string SqlData = "server=10.194.48.150\\MySQL;database=FCT_LED_KanBan;uid=sa;pwd=Aa123456;Connection Timeout=3";

        public static DataSet GetDataSet(string strSql)
        {
            SqlConnection cn = new SqlConnection(SqlData);
            cn.Open();
            SqlCommand cmd = new SqlCommand(strSql, cn);
            SqlDataAdapter dp = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            dp.Fill(ds);
            cn.Close();
            return ds;
        }

    }
}
