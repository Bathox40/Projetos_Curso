using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model_VO
{
    public class Pedidos_Interior_VO
    {
        private int id;
        private Cliente_VO cliente_id;
        private string descricao;
        private int estado;

        public Pedidos_Interior_VO()
        {

        }
        public Pedidos_Interior_VO(int intID, Cliente_VO CCliente, string strDescricao, int intEstado)
        {
            ID = intID;
            Cliente_ID = CCliente;
            Descricao = strDescricao;
            Estado = intEstado;
        }


        public int ID
        {
            get => id;
            set => id = value;
        }

        public Cliente_VO Cliente_ID
        {
            get => cliente_id;
            set => cliente_id = value;
        }

        public string Descricao
        {
            get => descricao;
            set => descricao = value;
        }

        public int Estado
        {
            get => estado;
            set => estado = value;
        }
    }
}
