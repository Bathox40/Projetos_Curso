using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Model_VO
{
    public class Cliente_VO
    {
        private int id;
        private string nome;
        private string descricao;
        private int ativos;

        public Cliente_VO()
        {

        }
        public Cliente_VO(int intID, string strnome, string strDescricao, int intAtivos)
        {
            ID = intID;
            Nome = strnome;
            Descricao = strDescricao;
            Ativos = intAtivos;
        }

        public int ID
        {
            get { return this.id; }
            set { this.id = value; }
        }

        public string Nome
        {
            get { return this.nome; }
            set { this.nome = value; }
        }
        public string Descricao
        {
            get { return this.descricao; }
            set { this.descricao = value; }
        }
        public int Ativos
        {
            get { return this.ativos; }
            set { this.ativos = value; }
        }
    }
}
