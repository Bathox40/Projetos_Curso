using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Camada_DAO_DAL;
using Model_VO;

namespace Facade_FD
{
    public  class Pedidos_Exterior_FD
    {
        Pedidos_Exterior_DAO objPedidos_Exterior_DAO;

        public DataTable ConsultarBD(Pedidos_Exterior_VO objvo_VO)
        {
            try
            {
                objPedidos_Exterior_DAO = new Pedidos_Exterior_DAO();
                return objPedidos_Exterior_DAO.ConsultarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool IncluirBD(Pedidos_Exterior_VO objvo_VO)
        {
            try
            {
                objPedidos_Exterior_DAO = new Pedidos_Exterior_DAO();
                return objPedidos_Exterior_DAO.IncluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool ExcluirBD(Pedidos_Exterior_VO objvo_VO)
        {
            try
            {
                objPedidos_Exterior_DAO = new Pedidos_Exterior_DAO();
                return objPedidos_Exterior_DAO.ExcluirBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool AlterarBD(Pedidos_Exterior_VO objvo_VO)
        {
            try
            {
                objPedidos_Exterior_DAO = new Pedidos_Exterior_DAO();
                return objPedidos_Exterior_DAO.AlterarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
