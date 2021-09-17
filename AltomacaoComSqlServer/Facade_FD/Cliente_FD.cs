using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Model_VO;
using Camada_DAO_DAL;

namespace Facade_FD
{
    public class Cliente_FD
    {

        Cliente_DAO objCliente_DAO;

        public List<string> ImportarBDC()
        {
            try
            {
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.ImportarBDC();
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.ImportarBDD();
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.ConsultarBD(objvo_VO);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.IncluirBD(objvo_VO);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.ExcluirBD(objvo_VO);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.AlterarBD(objvo_VO);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.Consultar_Clientes_Sem_Pedidos_Exterior();
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.Consultar_Pedidos_De_Clientes_Slecionados(strPesquisarIdClientes);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.Consultar_De_Quantidade_De_PedidosInterior_Por_Clientes(strPesquisarIdClientes);
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
                objCliente_DAO = new Cliente_DAO();
                return objCliente_DAO.Consultar_Quantidades_De_Pedidos_Dos_Clientes();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
