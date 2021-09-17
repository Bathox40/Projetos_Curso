using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Model_VO;
using Camada_DAO_DAL;


namespace Facade_FD
{
    public class Pedidos_Interior_FD
    {
        Pedidos_Interior_DAO objPedidos_Interior_DAO;

        public DataTable ConsultarBD(Pedidos_Interior_VO objvo_VO)
        {
            try
            {
                objPedidos_Interior_DAO = new Pedidos_Interior_DAO();
                return objPedidos_Interior_DAO.ConsultarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBD(Pedidos_Interior_VO objvo_VO)
        {
            try
            {
                objPedidos_Interior_DAO = new Pedidos_Interior_DAO();
                return objPedidos_Interior_DAO.IncluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBD(Pedidos_Interior_VO objvo_VO)
        {
            try
            {
                objPedidos_Interior_DAO = new Pedidos_Interior_DAO();
                return objPedidos_Interior_DAO.ExcluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBD(Pedidos_Interior_VO objvo_VO)
        {
            try
            {
                objPedidos_Interior_DAO = new Pedidos_Interior_DAO();
                return objPedidos_Interior_DAO.AlterarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
