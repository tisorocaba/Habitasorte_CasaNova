using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Habitasorte.Business.Model.Publicacao {
    public class CandidatoPub {

        public int? IdCandidato { get; set; }
        public string Cpf { get; set; }
        public string Nome { get; set; }
        public int IdInscricao { get; set; }

        [JsonIgnore]
        public int QuantidadeCriterios { get; set;  }
    }
}
