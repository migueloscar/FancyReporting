using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using FancyReporting.Models;
using System.Collections.Generic;
using System;
using System.Linq;

namespace FancyReporting
{
    public static class RepGenGraficoEstadistico
    {
        public static Paragraph genParrafoGrafico(string cod, string numChart)
        {
            return HerGenGraficoEstadistico.generarParrafoGrafico(cod, numChart);
        }
        public static void generarGraficoEstadistico(MainDocumentPart mainDocumentPart,
                                                    string codigo, 
                                                    List<DatoPreparado> datosGrafico,
                                                    string tituloGraEsta,
                                                    Dictionary<string,string> cadenas)
        {
            HerGenGraficoEstadistico.generarGraficoEstadistico(mainDocumentPart,codigo,datosGrafico, tituloGraEsta,cadenas);
        }
        public static string generarTituloGrafico(Dictionary<string, string> cadenas, List<DateTime> periodos, List<DatoPreparado> datos)
        {
            string formatoFechas = cadenas["formatoFecha"];
            string dato = cadenas["dato"];

            string res = cadenas["inicioTitGraEst"];
            res += (datos.Count == 1 ? cadenas["artSingular"] : cadenas["artPlural"]) + " ";
            res += (datos.Count == 1 ? dato : RepGenParrafos.pluralizarPalabra(dato)) + " ";
            res += detalleListadoSoloTexto(cadenas, datos.Select(x => x.detalle).ToList()) + " ";
            res += string.Format(cadenas["entreFechasFormato1"],
                                 periodos.First().ToString(formatoFechas),
                                 periodos.Last().ToString(formatoFechas)
                                 );
            return res.ToUpper();
        }

        public static string detalleListadoSoloTexto(Dictionary<string, string> cadenas, List<string> listaDetalles)
        {
            string cadena = "";
            if (listaDetalles.Count == 1)
            {
                cadena = listaDetalles[0];
            }
            else
            {
                cadena += String.Join(", ", listaDetalles.ToArray(), 0, listaDetalles.Count - 1);
                cadena += " "+cadenas["ultimoConector"]+" " + listaDetalles.Last();
            }

            return cadena;
        }
        public static Paragraph genIntroGraEst(Dictionary<string,string> cadenas, int numElementos, string grupo) 
        {
            string intro = "";
            if (numElementos > 1)
            {
                intro = string.Format(cadenas["intGraEstPlural"], cadenas["artPlural"], RepGenParrafos.pluralizarPalabra(cadenas["dato"]), grupo); 
            }
            else
            {
                intro = string.Format(cadenas["intGraEstSingular"], cadenas["artSingular"], cadenas["dato"], grupo);
            }
            
            return HerGenParrafo.agregarParrafo(intro);
        }
    }
}
