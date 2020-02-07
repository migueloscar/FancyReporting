using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FancyReporting.Models
{
    public class PartesInforme
    {
        public int modeloInforme { get; set; }
        public Dictionary<string, string> cadenas { get; set; }
        public List<RepGenTabla.resultado> listaTabColResIntros { get; set; }
        public List<RepGenParrafos.generarCriterio> listaParrafosCriterios { get; set; }
        public List<RepGenParrafos.generarIntro> listaParrafosIntros { get; set; }
    }
}
