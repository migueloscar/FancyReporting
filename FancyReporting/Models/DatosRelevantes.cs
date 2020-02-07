using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FancyReporting.Models
{
    public class DatosRelevantes
    {
        public int orden { get; set; }
        public string detalle { get; set; }
        public double comparacion { get; set; }
        public int tendecia { get; set; }
    }
}