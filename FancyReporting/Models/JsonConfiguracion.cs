using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FancyReporting.Models
{
    public class JsonConfiguracion
    {
        public string model { get;set; }
        public List<DatosParametrosConfiguracion> datos { get; set; }
        public List<GrupoDato> grupos { get; set; }
        public Dictionary<string, string> cadenas { set; get; }
    }
}