using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using FancyReporting.Models;
namespace FancyReporting
{
    public static class RepGenParrafos
    {
        public delegate string generarIntro(List<DateTime> periodos);
        public delegate int generarCriterio(List<PeriodoValor> periodos);

        public static Paragraph generarDescripcion(
            Dictionary<string, string> cadenas,
            List<DateTime> periodos, 
            List<DatosRelevantes> listDatos,
            int iParrafo, 
            string intro,
            string[] listConectores)
        {
            string dato = cadenas["dato"];

            //obteniendo por grupos las que incrementaron y decrementaron
            int datosAscienden = listDatos.Where(x => x.tendecia > 0).Count();
            int datosDescienden = listDatos.Where(x => x.tendecia < 0).Count();

            List<Run> descripcionTabla = new List<Run>();
            descripcionTabla.Add(HerGenParrafo.addTextNormal(intro));
            descripcionTabla.Add(HerGenParrafo.agregarEspacio());

            if (datosAscienden != 0)
            {
                if (datosAscienden == 1)
                {
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["artSingular"] + " " + dato));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(detallesListado(cadenas, listDatos.Where(x => x.tendecia > 0).Select(x => x.detalle).ToList()));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(listConectores[0]));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(resultadosListado(cadenas, listDatos.Where(x => x.tendecia > 0).Select(x => x.comparacion).ToList()));
                }
                else
                {
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["artPlural"] + " " + pluralizarPalabra(dato)));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(detallesListado(cadenas, listDatos.Where(x => x.tendecia > 0).Select(x => x.detalle).ToList()));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(listConectores[1]));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(resultadosListado(cadenas, listDatos.Where(x => x.tendecia > 0).Select(x => x.comparacion).ToList()));
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["finDescripcionPlural"]));
                }
                if (datosDescienden != 0)
                {
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["contraste"]));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                }
            }
            if (datosDescienden != 0)
            {
                if (datosDescienden == 1)
                {
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["artSingular"] + " " + dato));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(detallesListado(cadenas, listDatos.Where(x => x.tendecia < 0).Select(x => x.detalle).ToList()));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(listConectores[2]));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(resultadosListado(cadenas, listDatos.Where(x => x.tendecia < 0).Select(x => x.comparacion).ToList()));
                }
                else
                {
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["artPlural"] + " " + pluralizarPalabra(dato)));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(detallesListado(cadenas, listDatos.Where(x => x.tendecia < 0).Select(x => x.detalle).ToList()));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(listConectores[3]));
                    descripcionTabla.Add(HerGenParrafo.agregarEspacio());
                    descripcionTabla.AddRange(resultadosListado(cadenas, listDatos.Where(x => x.tendecia < 0).Select(x => x.comparacion).ToList()));
                    descripcionTabla.Add(HerGenParrafo.addTextNormal(cadenas["finDescripcionPlural"]));


                }
                descripcionTabla.Add(HerGenParrafo.addTextNormal("."));
            }
            else
            {
                descripcionTabla.Add(HerGenParrafo.addTextNormal("."));
            }

            Paragraph paragraph = HerGenParrafo.agregarParrafo(descripcionTabla);

            return paragraph;
        }

        public static string introGraficoEstadistico(Dictionary<string, string> cadenas, List<DatosRelevantes> datosInforme, string dato)
        {
            string res = "";
            if (datosInforme.Count() == 1)
            {
                res = string.Format(cadenas["intGraEstSingular"], dato);
            }
            else
            {
                res = string.Format(cadenas["intGraEstPlural"], pluralizarPalabra(dato));
            }
            return res;
        }

        public static List<Run> detallesListado(Dictionary<string, string> cadenas, List<string> datosInforme)
        {
            List<Run> listRun = new List<Run>();
            if (datosInforme.Count == 1)
            {
                listRun.Add(HerGenParrafo.addTextBold(datosInforme[0]));
            }
            else if (datosInforme.Count == 2)
            {
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.First()));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextNormal(cadenas["ultimoConector"]));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.Last()));
            }
            else
            {
                int numEle = datosInforme.Count - 2;
                for(int i=0; i< numEle; i++)
                {
                    listRun.Add(HerGenParrafo.addTextBold(datosInforme[i]));
                    listRun.Add(HerGenParrafo.addTextNormal(", "));
                }
                listRun.Add(HerGenParrafo.addTextBold(datosInforme[numEle]));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextNormal(cadenas["ultimoConector"]));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.Last()));
            }
            return listRun;
        }

        public static List<Run> resultadosListado(Dictionary<string, string> cadenas, List<double> datosInforme)
        {
            List<Run> listRun = new List<Run>();
            if (datosInforme.Count == 1)
            {
                listRun.Add(HerGenParrafo.addTextBold(datosInforme[0].ToString(cadenas["formatoParrResultado"])));
            }
            else if(datosInforme.Count == 2)
            {
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.First().ToString(cadenas["formatoParrResultado"])));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextNormal(cadenas["ultimoConector"]));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.Last().ToString(cadenas["formatoParrResultado"])));
            }
            else
            {
                int numEle = datosInforme.Count - 2;
                for (int i = 0; i < numEle; i++)
                {
                    listRun.Add(HerGenParrafo.addTextBold(datosInforme[i].ToString(cadenas["formatoParrResultado"])));
                    listRun.Add(HerGenParrafo.addTextNormal(", "));
                }
                listRun.Add(HerGenParrafo.addTextBold(datosInforme[numEle].ToString(cadenas["formatoParrResultado"])));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextNormal(cadenas["ultimoConector"]));
                listRun.Add(HerGenParrafo.agregarEspacio());
                listRun.Add(HerGenParrafo.addTextBold(datosInforme.Last().ToString(cadenas["formatoParrResultado"])));
            }
            return listRun;
        }

        public static Paragraph agregarSubTitulo(string cadena)
        {
            return HerGenParrafo.agregarSubTitulo(cadena);
        }

        public static string pluralizarPalabra(string cadena)
        {
            return cadena + (Regex.IsMatch(cadena, @"([aeiou])$") ? "s" : "es");
        }

        public static Paragraph agregarSubTituloGeneral(string cadena)
        {
            return HerGenParrafo.agregarSubTituloGeneral(cadena.ToUpper());
        }


    }
}
