using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.IO;
using Model_VO;
using Facade_FD;


namespace Camada_BLL
{
    public class Cliente_BLL
    {

        Cliente_FD objCliente_FD;

        StreamReader objLeitor;
        string strLinhaLida;

        public List<string> ImportarTXT()
        {
            try
            {
                List<string> resultado = new List<string>();
                objLeitor = new StreamReader(@"C:\Atomacao_Bancaria\Clientes.txt");
                strLinhaLida = objLeitor.ReadLine();

                while (strLinhaLida != null)
                {
                    resultado.Add(strLinhaLida);
                    strLinhaLida = objLeitor.ReadLine();
                }
                return resultado;
            }
            catch (Exception ex)
            {

                throw new Exception("Falha no Importar Conectado : " + ex.Message);
            }

        }

        public List<string> ImportarBDC()
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.ImportarBDC();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public List<string> ImportarBDD()
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.ImportarBDD();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable ConsultarBD(Cliente_VO objvo_VO)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.ConsultarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBD(Cliente_VO objvo_VO)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.IncluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBD(Cliente_VO objvo_VO)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.ExcluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBD(Cliente_VO objvo_VO)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.AlterarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable Consultar_Clientes_Sem_Pedidos_Exterior()
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.Consultar_Clientes_Sem_Pedidos_Exterior();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable Consultar_Pedidos_De_Clientes_Slecionados(string strPesquisarIdClientes)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.Consultar_Pedidos_De_Clientes_Slecionados(strPesquisarIdClientes);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable Consultar_De_Quantidade_De_PedidosInterior_Por_Clientes(string strPesquisarIdClientes)
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.Consultar_De_Quantidade_De_PedidosInterior_Por_Clientes(strPesquisarIdClientes);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable Consultar_Quantidades_De_Pedidos_Dos_Clientes()
        {
            try
            {
                objCliente_FD = new Cliente_FD();
                return objCliente_FD.Consultar_Quantidades_De_Pedidos_Dos_Clientes();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
