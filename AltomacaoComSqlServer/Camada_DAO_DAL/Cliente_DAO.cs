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
    public class Cliente_DAO : DAO_DAO
    {
        OleDbCommand objComando;
        OleDbDataAdapter objAdaptador;
        OleDbDataReader objLeitorBD;
        DataTable objTabela;

        public List<string> ImportarBDC()
        {
            try
            {
                List<string> resultado = new List<string>();
                OpenCon();
                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT");
                strSql.Append(" Nome");
                strSql.Append(" FROM");
                strSql.Append(" Clientes");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objLeitorBD = objComando.ExecuteReader();

                while (objLeitorBD.Read())
                {
                    resultado.Add(objLeitorBD["Nome"].ToString());
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Importar Conectado : " + ex.Message);
            }
            finally
            {
                CloseCon();
            }
        }

        public List<string> ImportarBDD()
        {
            try
            {
                List<string> resultado = new List<string>();


                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT");
                strSql.Append(" Nome");
                strSql.Append(" FROM");
                strSql.Append(" Clientes");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                foreach (DataRow LinhaDaTabela in objTabela.Rows)
                {
                    resultado.Add(LinhaDaTabela["Nome"].ToString());
                }

                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Importar Desconectado : " + ex.Message);
            }

        }

        public override DataTable ConsultarBD(Object objvo_VO)
        {
            try
            {
                Cliente_VO objParCliente_VO = (Cliente_VO)objvo_VO;

                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT");
                strSql.Append(" ID");
                strSql.Append(",Nome ");
                strSql.Append(",Descricao ");
                strSql.Append(",Ativos ");
                strSql.Append(" FROM");
                strSql.Append(" Clientes");


                if (!objParCliente_VO.ID.Equals(0))
                {
                    strSql.Append(" WHERE");
                    strSql.Append(" ID = ?");

                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                    objComando.Parameters.Add("?ID", OleDbType.BigInt);
                    objComando.Parameters["?ID"].Value = objParCliente_VO.ID;
                }
                else if (!string.IsNullOrEmpty(objParCliente_VO.Nome))
                {
                    strSql.Append(" WHERE");
                    strSql.Append(" Nome = ?");

                    objComando = new OleDbCommand(strSql.ToString(), getConexao());
                    objComando.Parameters.Add("?Nome", OleDbType.VarChar);
                    objComando.Parameters["?Nome"].Value = objParCliente_VO.Nome;
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
                Cliente_VO objParCliente_VO = (Cliente_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("INSERT");
                strSql.Append(" INTO");
                strSql.Append(" Clientes (");
                strSql.Append(" Nome ");
                strSql.Append(",Descricao ");
                strSql.Append(",Ativos");
                strSql.Append(" ) VALUES (");
                strSql.Append(" ?");
                strSql.Append(",? ");
                strSql.Append(",?)");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?Nome", OleDbType.VarChar);
                objComando.Parameters["?Nome"].Value = objParCliente_VO.Nome;

                objComando.Parameters.Add("?Descricao", OleDbType.VarChar);
                objComando.Parameters["?Descricao"].Value = objParCliente_VO.Descricao;

                objComando.Parameters.Add("?Ativos", OleDbType.BigInt);
                objComando.Parameters["?Ativos"].Value = objParCliente_VO.Ativos;

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
                Cliente_VO objParCliente_VO = (Cliente_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("DELETE");
                strSql.Append(" FROM");
                strSql.Append(" Clientes");
                strSql.Append(" WHERE");
                strSql.Append(" ID = ? ");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?ID", OleDbType.BigInt);
                objComando.Parameters["?ID"].Value = objParCliente_VO.ID;

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
                Cliente_VO objParCliente_VO = (Cliente_VO)objvo_VO;

                OpenCon();

                StringBuilder strSql = new StringBuilder();
                strSql.Append("UPDATE");
                strSql.Append(" Clientes");
                strSql.Append(" SET");
                strSql.Append(" Nome = ? ");
                strSql.Append(",Descricao = ?");
                strSql.Append(",Ativos = ? ");
                strSql.Append(" WHERE");
                strSql.Append(" ID = ? ");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());
                objComando.Parameters.Add("?Nome", OleDbType.VarChar);
                objComando.Parameters["?Nome"].Value = objParCliente_VO.Nome;

                objComando.Parameters.Add("?Descricao", OleDbType.VarChar);
                objComando.Parameters["?Descricao"].Value = objParCliente_VO.Descricao;

                objComando.Parameters.Add("?Ativos", OleDbType.SmallInt);
                objComando.Parameters["?Ativos"].Value = objParCliente_VO.Ativos;

                objComando.Parameters.Add("?ID", OleDbType.BigInt);
                objComando.Parameters["?ID"].Value = objParCliente_VO.ID;

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
        public DataTable Consultar_Clientes_Sem_Pedidos_Exterior()
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT ID, Nome, Descricao, Ativos");
                strSql.Append(" FROM Clientes");
                strSql.Append(" SELECT C.ID, C.Nome, C.Descricao, C.Ativos");
                strSql.Append(" FROM Cliente AS C");
                strSql.Append(" INNER JOIN Pedidos_Exterior PE");
                strSql.Append(" ON C.ID = PE.Cliente_ID");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }

            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar BD : " + ex.Message);
            }
        }

        public DataTable Consultar_Pedidos_De_Clientes_Slecionados(string strPesquisarIdClientes)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT C.ID, C.Nome, C.Descricao, P.ID AS Cliente_ID, P.Descricao, P.Estado ");
                strSql.Append(" FROM Clientes C,");
                strSql.Append(" (SELECT I.Cliente_ID, I.ID, I.Descricao, I.Estado");
                strSql.Append(" FROM Pedidos_Interior I");
                strSql.Append(" UNION");
                strSql.Append(" SELECT E.Cliente_ID, E.ID, E.Descricao, E.Estado");
                strSql.Append(" FROM Pedidos_Exterior E) P");
                strSql.Append(" WHERE C.ID IN(" + strPesquisarIdClientes + ") AND C.ID = P.Cliente_ID");
                strSql.Append(" ORDER BY C.ID, P.Cliente_ID");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }

            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar BD : " + ex.Message);
            }
        }

        public DataTable Consultar_De_Quantidade_De_PedidosInterior_Por_Clientes(string strPesquisarIdClientes)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT COUNT (ID) AS Quantidades");
                strSql.Append(" FROM Pedidos_Interior");
                strSql.Append(" WHERE Cliente_ID IN (" + strPesquisarIdClientes + ")");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }

            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar BD : " + ex.Message);
            }
        }

        public DataTable Consultar_Quantidades_De_Pedidos_Dos_Clientes()
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                strSql.Append("SELECT ID AS ID_DO_CLIENTE, Nome AS NOME, QUANTIDADE_PEDIDOS_POR_CLIENTE AS QUANTIDADE_TOTAL FROM");
                strSql.Append(" (SELECT Convert(varchar, ID) AS ID, Nome, COUNT(*) AS QUANTIDADE_PEDIDOS_POR_CLIENTE FROM");
                strSql.Append(" (SELECT C.ID, C.Nome, PIN.ID AS ID_PEDIDO, PIN.Descricao FROM Clientes C ");
                strSql.Append(" INNER JOIN Pedidos_Interior PIN ON C.ID = PIN.Cliente_ID");
                strSql.Append(" UNION ALL");
                strSql.Append(" SELECT C.ID, C.Nome, PE.ID AS ID_PEDIDO, PE.Descricao FROM Clientes C");
                strSql.Append(" INNER JOIN Pedidos_Exterior PE ON C.ID = PE.Cliente_ID) AS TABELAS GROUP BY ID, Nome");
                strSql.Append(" UNION");
                strSql.Append(" SELECT 'XXXXXXXXXXX', 'TOTAL GERAL ===>', SUM(QUANTIDADE_PEDIDOS_POR_CLIENTE) AS QUANTITIDADE_TOTAL FROM");
                strSql.Append(" (SELECT ID, Nome, COUNT(*) AS QUANTIDADE_PEDIDOS_POR_CLIENTE");
                strSql.Append(" FROM");
                strSql.Append(" (SELECT C.ID, C.Nome, PIN.ID AS ID_PEDIDO, PIN.Descricao FROM Clientes C");
                strSql.Append(" INNER JOIN Pedidos_Interior PIN ON C.ID = PIN.Cliente_ID");
                strSql.Append(" UNION ALL");
                strSql.Append(" SELECT C.ID, C.Nome, PE.ID AS ID_PEDIDO, PE.Descricao FROM Clientes C");
                strSql.Append(" INNER JOIN Pedidos_Exterior PE ON C.ID = PE.Cliente_ID) AS TABELAS GROUP BY ID, Nome) TOTAL) AS GERAL");

                objComando = new OleDbCommand(strSql.ToString(), getConexao());

                objAdaptador = new OleDbDataAdapter(objComando);
                objTabela = new DataTable();
                objAdaptador.Fill(objTabela);

                return objTabela;
            }

            catch (Exception ex)
            {
                throw new Exception("Falha no Alterar BD : " + ex.Message);
            }
        }

    }
}
