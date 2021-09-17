using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Model_VO;

namespace Camada_DAO_DAL
{
    public class Pedidos_Exterior_DAO : DAO_DAO
    {

        Cliente_VO objCliente_VO;
        OleDbCommand objComando;
        OleDbDataAdapter objAdaptador;
        DataTable objTabela;

        public override DataTable ConsultarBD(Object objvo_VO)
        {
            try
            {
                Pedidos_Exterior_VO objParPedidos_Exterior_VO = (Pedidos_Exterior_VO)objvo_VO;

                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT");
                strSql.Append(" ID");
                strSql.Append(",Cliente_ID ");
                strSql.Append(",Descricao ");
                strSql.Append(",Estado ");
                strSql.Append(" FROM");
                strSql.Append(" Pedidos_Exterior");


                if (!objParPedidos_Exterior_VO.ID.Equals(0))
                {
                    strSql.Append(" WHERE");
                    strSql.Append(" ID = ?");

                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                    objComando.Parameters.Add("?ID", OleDbType.BigInt);
                    objComando.Parameters["?ID"].Value = objParPedidos_Exterior_VO.ID;
                }
                else if (!string.IsNullOrEmpty(objParPedidos_Exterior_VO.Descricao))
                {
                    strSql.Append(" WHERE");
                    strSql.Append(" Descricao = ?");

                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                    objComando.Parameters.Add("?Descricao", OleDbType.VarChar);
                    objComando.Parameters["?Descricao"].Value = objParPedidos_Exterior_VO.Descricao;
                }
                else if (!objParPedidos_Exterior_VO.Cliente_ID.ID.Equals(0))
                {
                    strSql.Append(" WHERE");
                    strSql.Append(" Cliente_ID = ?");

                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                    objComando.Parameters.Add("?Cliente_ID", OleDbType.BigInt);
                    objComando.Parameters["?Cliente_ID"].Value = objParPedidos_Exterior_VO.Cliente_ID.ID;
                }
                else
                {
                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                }

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Consultar BD : " + ex.Message);
            }

        }

        public override bool IncluirBD(Object objvo_VO)
        {
            try
            {
                Pedidos_Exterior_VO objParPedidos_Exterior_VO = (Pedidos_Exterior_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("INSERT");
                strSql.Append(" INTO");
                strSql.Append(" Pedidos_Exterior (");
                strSql.Append(" Cliente_ID ");
                strSql.Append(",Descricao ");
                strSql.Append(",Estado");
                strSql.Append(" ) VALUES (");
                strSql.Append(" ?");
                strSql.Append(",? ");
                strSql.Append(",?)");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?Cliente_ID", OleDbType.BigInt);
                objComando.Parameters["?Cliente_ID"].Value = objParPedidos_Exterior_VO.Cliente_ID.ID;

                objComando.Parameters.Add("?Descricao", OleDbType.VarChar);
                objComando.Parameters["?Descricao"].Value = objParPedidos_Exterior_VO.Descricao;

                objComando.Parameters.Add("?Estado", OleDbType.SmallInt);
                objComando.Parameters["?Estado"].Value = objParPedidos_Exterior_VO.ID;

                if (objComando.ExecuteNonQuery() > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Incluir BD : " + ex.Message);
            }
            finally
            {
                CloseCon();
            }
        }

        public override bool ExcluirBD(Object objvo_VO)
        {
            try
            {
                Pedidos_Exterior_VO objParPedidos_Exterior_VO = (Pedidos_Exterior_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("DELETE");
                strSql.Append(" FROM");
                strSql.Append(" Pedidos_Exterior");
                strSql.Append(" WHERE");
                strSql.Append(" ID = ? ");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?ID", OleDbType.BigInt);
                objComando.Parameters["?ID"].Value = objParPedidos_Exterior_VO.ID;

                if (objComando.ExecuteNonQuery() > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Excluir BD : " + ex.Message);
            }
            finally
            {
                CloseCon();
            }

        }

        public override bool AlterarBD(Object objvo_VO)
        {
            try
            {
                Pedidos_Exterior_VO objParPedidos_Exterior_VO = (Pedidos_Exterior_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("UPDATE");
                strSql.Append(" Pedidos_Exterior");
                strSql.Append(" SET");
                strSql.Append(" Cliente_ID = ? ");
                strSql.Append(",Descricao = ?");
                strSql.Append(",Estado = ? ");
                strSql.Append(" WHERE");
                strSql.Append(" ID");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?Cliente_ID", OleDbType.VarChar);
                objComando.Parameters["?Cliente_ID"].Value = objParPedidos_Exterior_VO.Cliente_ID.ID;

                objComando.Parameters.Add("?Descricao", OleDbType.VarChar);
                objComando.Parameters["?Descricao"].Value = objParPedidos_Exterior_VO.Descricao;

                objComando.Parameters.Add("?Estado", OleDbType.SmallInt);
                objComando.Parameters["?Estado"].Value = objParPedidos_Exterior_VO.Estado;

                objComando.Parameters.Add("?ID", OleDbType.BigInt);
                objComando.Parameters["?ID"].Value = objParPedidos_Exterior_VO.ID;

                if (objComando.ExecuteNonQuery() > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar BD : " + ex.Message);
            }
            finally
            {
                CloseCon();
            }
        }

    }
}
