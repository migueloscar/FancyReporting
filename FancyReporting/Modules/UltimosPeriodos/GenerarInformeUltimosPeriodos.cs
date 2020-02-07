using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FancyReporting.Models;
using System.Reflection;
using Newtonsoft.Json;

namespace FancyReporting.Modules.UltimosPeriodos
{
    public class GeneracionInformeUltimosPeriodos:Informe
    {

        public override void cargandoCadenasDefecto()
        {
            configuracion.cadenas = CadenasDefaultUltimosPeriodos.leerCadenas;
        }

        public override void agregandoValidaciones()
        {
            agregarFuncionValidacion(valDatRepetidos);
        }

        public override void agregandoFuncionesResultado()
        {
            agregarFuncionesColResultado(calcularTotal);
        }
        public override void agregandoTitColRes()
        {
            agregarColResultadoTitulo("Acumulado");
        }
        //funciones parrafos textos
        public override void agregandoIntrosParrafos()
        {
            agregarIntParrafo(iniParIntro);
        }
        public override void agregandoCriComentario()
        {
            agregarCriComentario(iniCriEva);
        }
        public override void agregandoConector()
        {
            agregarConector(new string[]{
                                configuracion.cadenas["asciendeSingular"], 
                                configuracion.cadenas["asciendePlural"],
                                configuracion.cadenas["desciendeSingular"], 
                                configuracion.cadenas["desciendePlural"]
            });
        }

        //columnas funciones
        public double calcularTotal(List<PeriodoValor> listaPeriodos)
        {
            return listaPeriodos.Sum(x => x.valor);
        }
       
        //introduccion parrafos
        public string iniParIntro(List<DateTime> periodos)
        {
            string intro = configuracion.cadenas["iniParIni"] + " ";
            intro += string.Format(configuracion.cadenas["entreFechasFormato1"],
                    periodos[0].ToString(configuracion.cadenas["formatoFecha"]),
                    periodos[1].ToString(configuracion.cadenas["formatoFecha"]));
            return intro;
        }
        
        //criterios
        public int iniCriEva(List<PeriodoValor> listaPeriodos)
        {
            return calcularTotal(listaPeriodos) < 0 ? -1 : 1;
        }

        public List<string> valDatRepetidos(List<JsonDatos> jsonDatos)
        {
            List<string> listMensaValidacion = new List<string>();
            var ld = jsonDatos.GroupBy(x => new { x.periodo, x.detalle }).Where(g => g.Count() > 1).Select(y => y.Key).ToList();
            List<JsonDatos> listDuplicados = (from dat in jsonDatos
                                              join datDup in ld
                                              on
                                                 new { dat.detalle, dat.periodo } equals new { datDup.detalle, datDup.periodo }
                                              select new JsonDatos() { detalle = dat.detalle, periodo = dat.periodo, valor = dat.valor }
                                            ).ToList();

            if (listDuplicados.Count() > 0)
            {
                List<string> cadenas = new List<string>();
                foreach (JsonDatos obj in listDuplicados)
                {
                    cadenas.Add(obj.detalle + " - " + obj.periodo.ToString(configuracion.cadenas["formatoFecha"]) + " - " + obj.valor);
                }

                string err = "Para proceder con el informe los datos deben ser unicos, sin embargo, la informacion obetenida tiene datos duplicados: \n" + string.Join("; ", cadenas);

                listMensaValidacion.Add(err);
            }

            return listMensaValidacion;
        }

    }
}
