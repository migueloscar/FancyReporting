using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FancyReporting.Models;

namespace FancyReporting.Modules.IVPeriodos
{
    public class GenerarInformeIVPeriodos:Informe
    {

        public override void cargandoCadenasDefecto()
        {
            configuracion.cadenas = CadenasDefaultIVPeriodos.leerCadenas;
        }

        public override void agregandoValidaciones()
        {
            agregarFuncionValidacion(valNumColumnas);
            agregarFuncionValidacion(valDatRepetidos);
        }

        public override void agregandoFuncionesResultado()
        {
            agregarFuncionesColResultado(calcularPeriodoI);
            agregarFuncionesColResultado(calcularPeriodoII);
            agregarFuncionesColResultado(calculoFinal);
        }
        public override void agregandoTitColRes()
        {
            string formatoFechas = configuracion.cadenas["formatoFecha"];
            agregarColResultadoTitulo(string.Format("{0} / {1}", listPeriodos[0].ToString(formatoFechas), listPeriodos[1].ToString(formatoFechas)));
            agregarColResultadoTitulo(string.Format("{0} / {1}", listPeriodos[2].ToString(formatoFechas), listPeriodos[3].ToString(formatoFechas)));
            agregarColResultadoTitulo(configuracion.cadenas["evolucionFinal"]);
        }
        //funciones parrafos textos
        public override void agregandoIntrosParrafos()
        {
            agregarIntParrafo(iniParIntro);
            agregarIntParrafo(medioParIntro);
            agregarIntParrafo(finalParIntro);
        }
        public override void agregandoCriComentario()
        {
            agregarCriComentario(iniCriEva);
            agregarCriComentario(midCriEva);
            agregarCriComentario(finCriEva);
        }
        public override void agregandoConector()
        {
            agregarConector(new string[]{
                                configuracion.cadenas["asciendeSingular"], configuracion.cadenas["asciendePlural"],
                                configuracion.cadenas["desciendeSingular"], configuracion.cadenas["desciendePlural"]
            });
            agregarConector(new string[]{
                                configuracion.cadenas["asciendeSingular"], configuracion.cadenas["asciendePlural"],
                                configuracion.cadenas["desciendeSingular"], configuracion.cadenas["desciendePlural"]
            });
            agregarConector(new string[]{
                                configuracion.cadenas["tenFinAsciendeSingular"], configuracion.cadenas["tenFinAsciendePlural"],
                                configuracion.cadenas["tenFinDesciendeSingular"], configuracion.cadenas["tenFinDesciendePlural"]});
        }

        //columnas funciones
        public double calcularPeriodoI(List<PeriodoValor> listaPeriodosVal)
        {
            if (listaPeriodosVal[0].valor != 0)
            {
                return (listaPeriodosVal[1].valor - listaPeriodosVal[0].valor) / listaPeriodosVal[0].valor;
            }
            else
            {
                return 0;
            }
        }
        public double calcularPeriodoII(List<PeriodoValor> listaPeriodosVal)
        {
            if (listaPeriodosVal[0].valor != 0)
            {
                return (listaPeriodosVal[3].valor - listaPeriodosVal[2].valor) / listaPeriodosVal[2].valor;
            }
            else
            {
                return 0;
            }
        }

        public double calculoFinal(List<PeriodoValor> listaPeriodosVal)
        {
            return -calcularPeriodoI(listaPeriodosVal) + calcularPeriodoII(listaPeriodosVal);
        }
        //introduccion parrafos
        public string iniParIntro(List<DateTime> periodos)
        {
            string intro = configuracion.cadenas["iniParIni"]+" ";
            intro += string.Format(configuracion.cadenas["entreFechasFormato1"],
                    periodos[0].ToString(configuracion.cadenas["formatoFecha"]),
                    periodos[1].ToString(configuracion.cadenas["formatoFecha"]));
            return intro;
        }
        public string medioParIntro(List<DateTime> periodos)
        {
            string intro = configuracion.cadenas["iniParMed"] + " ";
            intro += string.Format(configuracion.cadenas["entreFechasFormato1"],
                periodos[2].ToString(configuracion.cadenas["formatoFecha"]),
                periodos[3].ToString(configuracion.cadenas["formatoFecha"]));
            return intro;
        }
        public string finalParIntro(List<DateTime> periodos)
        {
            string intro = configuracion.cadenas["iniParConFinal"] + " ";
            intro += string.Format(configuracion.cadenas["parrafoResultado"],
                                        periodos[1].ToString("yyyy"),
                                        periodos[1].ToString(configuracion.cadenas["formatoFecha"]),
                                        periodos[3].ToString("yyyy"),
                                        periodos[3].ToString(configuracion.cadenas["formatoFecha"]));
            return intro;
        }
        //criterios
        public int iniCriEva(List<PeriodoValor> listaPeriodosVal)
        {
            return calcularPeriodoI(listaPeriodosVal) < 0 ? -1 : 1;
        }

        public int midCriEva(List<PeriodoValor> listaPeriodosVal)
        {
            return calcularPeriodoII(listaPeriodosVal) < 0 ? -1 : 1;
        }

        public int finCriEva(List<PeriodoValor> listaPeriodosVal)
        {
            return calculoFinal(listaPeriodosVal) < 0 ? -1 : 1;
        }

        public List<string> valNumColumnas(List<JsonDatos> listDatos)
        {
            List<string> listMensaValidacion = new List<string>();
            
            if (listDatos.GroupBy(x=>x.periodo).Select(g=>g.Key).Count() < 4)
            {
                listMensaValidacion.Add(
                    "Este reporte requiere 4 periodos para generarse correctamente, " +
                    "comuniquese con Sistemas para la verificacion de los datos.");
            }

            return listMensaValidacion;
        }
        /// <summary>
        /// periodos y detalle repetidos
        /// </summary>
        /// <param name="listDatos"></param>
        /// <returns></returns>
        public List<string> valDatRepetidos(List<JsonDatos> jsonDatos)
        {
            List<string> listMensaValidacion = new List<string>();
            var ld = jsonDatos.GroupBy(x => new  { x.periodo,x.detalle }).Where(g => g.Count()>1).Select(y=>y.Key).ToList();
            List<JsonDatos> listDuplicados = (from dat in jsonDatos
                                              join datDup in ld
                                              on 
                                                 new { dat.detalle, dat.periodo } equals new { datDup.detalle ,datDup.periodo}
                                              select new JsonDatos() { detalle = dat.detalle, periodo = dat.periodo, valor = dat.valor }
                                            ).ToList();

            if (listDuplicados.Count()> 0)
            {
                List<string> cadenas= new List<string>();
                foreach(JsonDatos obj in listDuplicados)
                {
                    cadenas.Add(obj.detalle + " - " + obj.periodo.ToString(configuracion.cadenas["formatoFecha"])+" - "+obj.valor);
                }

                string err = "Para proceder con el informe los datos deben ser unicos, sin embargo, la informacion obetenida tiene datos duplicados: \n"+string.Join("; ",cadenas);

                listMensaValidacion.Add(err);
            }

            return listMensaValidacion;
        }
    }
}
