using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Camada_DAO_DAL
{
    public abstract class DAO_DAO : DB_DAO
    {

        public abstract DataTable ConsultarBD(Object objvo_VO);
        public abstract bool IncluirBD(Object objvo_VO);
        public abstract bool ExcluirBD(Object objvo_VO);
        public abstract bool AlterarBD(Object objvo_VO);

    }
}
