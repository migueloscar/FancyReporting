using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FancyReporting.Models
{
    public class DatosParametrosConfiguracion
    {
        public int orden { get; set; }
        public string detalle { get; set; }
        public int grupo { get; set; }
        public bool comentario { get; set; }
        public bool graficoEstadistico { get; set; }
    }
}