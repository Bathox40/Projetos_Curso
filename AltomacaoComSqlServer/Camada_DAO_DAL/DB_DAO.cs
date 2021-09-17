using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Configuration;

namespace Camada_DAO_DAL
{
    public class DB_DAO
    {  
            private static OleDbConnection objCon;

            public static OleDbConnection getConexao()
            {
                if (objCon == null)
                {
                    setConexao();
                }
                return objCon;
            }

            public static void setConexao()
            {
                objCon = new OleDbConnection(ConfigurationSettings.AppSettings["stringconexao"].ToString());
            }

            public static void OpenCon()
            {
                if (getConexao().State == System.Data.ConnectionState.Closed)
                {
                    objCon.Open();
                }
            }

            public static void CloseCon()
            {
                if (getConexao().State == System.Data.ConnectionState.Open)
                {
                    objCon.Close();
                    objCon.Dispose();
                    objCon = null;
                }
            }
    }
}
