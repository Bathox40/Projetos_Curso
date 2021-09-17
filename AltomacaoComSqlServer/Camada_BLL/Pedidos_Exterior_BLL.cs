using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Facade_FD;
using Model_VO;

namespace Camada_BLL
{
    public class Pedidos_Exterior_BLL
    {
        Pedidos_Exterior_FD objPedidos_Exterior_FD;

        public DataTable ConsultarBD(Pedidos_Exterior_VO objvo_VO)
        {
            try
            {
                objPedidos_Exterior_FD = new Pedidos_Exterior_FD();
                return objPedidos_Exterior_FD.ConsultarBD(objvo_VO);
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
                objPedidos_Exterior_FD = new Pedidos_Exterior_FD();
                return objPedidos_Exterior_FD.IncluirBD(objvo_VO);
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
                objPedidos_Exterior_FD = new Pedidos_Exterior_FD();
                return objPedidos_Exterior_FD.ExcluirBD(objvo_VO);
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
                objPedidos_Exterior_FD = new Pedidos_Exterior_FD();
                return objPedidos_Exterior_FD.AlterarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
