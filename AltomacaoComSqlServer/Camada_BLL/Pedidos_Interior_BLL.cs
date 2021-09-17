using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Model_VO;
using Facade_FD;

namespace Camada_BLL
{
    public class Pedidos_Interior_BLL
    {

        Pedidos_Interior_FD objPedidos_Interior_FD;

        public DataTable ConsultarBD(Pedidos_Interior_VO objvo_VO)
        {
            try
            {
                objPedidos_Interior_FD = new Pedidos_Interior_FD();
                return objPedidos_Interior_FD.ConsultarBD(objvo_VO);
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
                objPedidos_Interior_FD = new Pedidos_Interior_FD();
                return objPedidos_Interior_FD.IncluirBD(objvo_VO);
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
                objPedidos_Interior_FD = new Pedidos_Interior_FD();
                return objPedidos_Interior_FD.ExcluirBD(objvo_VO);
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
                objPedidos_Interior_FD = new Pedidos_Interior_FD();
                return objPedidos_Interior_FD.AlterarBD(objvo_VO);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
