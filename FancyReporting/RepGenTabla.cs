using FancyReporting.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FancyReporting
{
    public class RepGenTabla
    {
        public delegate List<string> validacion(List<JsonDatos> jsonDatos);
        public delegate double resultado(List<PeriodoValor> periodos);
        public static string[,] generarArrayTabla(
            JsonConfiguracion configuracion,
            List<resultado> ListFunciones,
            List<string> titulosResultados,
            List<DatoPreparado> datosPrep,
            List<DateTime> periodos,
            List<string> detalles,
            out Dictionary<string, string[]> tablaEstilos
            )
        {
            //formato primera columna
            List<string> anchos = new List<string>();
            List<string> alineaciones = new List<string>();
            List<string> columnasNegrita = new List<string>();
            List<string> filGrupos = new List<string>();

            string rubro = configuracion.cadenas["dato"];
            string formatoFechas = configuracion.cadenas["formatoFecha"];

            //inicio generacion tabla
            int filasCabezera = 2;
            int columnaDetalle = 1;

            int numArreglosContener = detalles.Count + filasCabezera + configuracion.grupos.Count;//o numero de filas
            int numElementosArreglo = columnaDetalle + periodos.Count + ListFunciones.Count;//o numero de columnas

            //generar tabla como array
            string[,] arrayDatos = new string[numArreglosContener, numElementosArreglo];

            //generarCabecera
            //primera fila
            int iFila = 0;
            int iColumna = 0;
            arrayDatos[iFila, iColumna] = RepGenParrafos.pluralizarPalabra(rubro).ToUpper();
            iColumna++;
            arrayDatos[iFila, iColumna] = configuracion.cadenas["tituloPeriodos"];
            iColumna++;
            for (int iPer = 0; iPer < periodos.Count - 1; iPer++)
            {
                arrayDatos[iFila, iColumna] = "MergeLeft";
                iColumna++;
            }
            if (ListFunciones.Count() > 0)
            {
                arrayDatos[iFila, iColumna] = configuracion.cadenas["tituloFunciones"];
                iColumna++;
                for (int iPer = 1; iPer < ListFunciones.Count; iPer++)
                {
                    arrayDatos[iFila, iColumna] = "MergeLeft";
                    iColumna++;
                }
            }
            //segunda fila
            iFila = 1;
            iColumna = 0;
            arrayDatos[iFila, iColumna] = "MergeUp";
            iColumna++;
            foreach (DateTime per in periodos)
            {
                arrayDatos[iFila, iColumna] = per.ToString(formatoFechas);
                iColumna++;
            }
            foreach (string titulos in titulosResultados)
            {
                arrayDatos[iFila, iColumna] = titulos;
                iColumna++;
            }

            //contenido tabla

            iFila++;
            iColumna = 0;

            //cuerpo
            configuracion.grupos = configuracion.grupos.OrderBy(x => x.grupo).ToList();
            foreach (GrupoDato gr in configuracion.grupos)
            {
                iColumna = 0;
                arrayDatos[iFila, iColumna] = gr.descripcion;
                iColumna++;
                int numColsPerYFun = periodos.Count + ListFunciones.Count;
                for (int icol = 0; icol < numColsPerYFun; icol++)
                {
                    arrayDatos[iFila, iColumna] = "MergeLeft";
                    iColumna++;
                }

                iFila++;
                filGrupos.Add(iFila.ToString());

                //obteniendo todos los detalles de un grupo
                List<string> detallesGrupo = configuracion.datos.Where(x => x.grupo.Equals(gr.grupo)).OrderBy(y => y.orden).ToList()//obtenemos todos los detalles de un grupo
                                            .Select(x => x.detalle).ToList();//obtenemos solo los nombres de los detalles.
                List<DatoPreparado> grupoListDatos = datosPrep.Where(x => detallesGrupo.Contains(x.detalle)).ToList();

                foreach (DatoPreparado dato in grupoListDatos)
                {
                    iColumna = 0;
                    arrayDatos[iFila, iColumna] = dato.detalle;
                    iColumna++;
                    foreach (PeriodoValor perVal in dato.periodos)
                    {
                        arrayDatos[iFila, iColumna] = perVal.valor.ToString(configuracion.cadenas["formatoTblValPeriodo"]);
                        iColumna++;
                    }
                    foreach (resultado funcion in ListFunciones)
                    {
                        arrayDatos[iFila, iColumna] = funcion(dato.periodos).ToString(configuracion.cadenas["formatoTblResultado"]);
                        iColumna++;
                    }
                    iFila++;
                }
            }

            Dictionary<string, string[]> tblEstilos = new Dictionary<string, string[]>();

            int numColumnas = periodos.Count + ListFunciones.Count + 1;

            //tblEstilos.Add("Alineaciones", alineaciones.ToArray());
            tblEstilos.Add("TamañoLetrasTabla", new string[] { "8" });

            int colIniFun = 1 + periodos.Count;
            for (int i = 1; i <= ListFunciones.Count; i++)
            {
                columnasNegrita.Add((colIniFun + i).ToString());
            }

            tblEstilos.Add("ColumnasNegrita", columnasNegrita.ToArray());
            tblEstilos.Add("FilasNegrita", filGrupos.ToArray());
            tblEstilos.Add("FilasGrupo", filGrupos.ToArray());
            tblEstilos.Add("celdaFinalCabecera", new string[] { "2", numColumnas.ToString() });
            //tablaEstilos.Add("ColumnasSubrayadas", new string[] { "6", "7", "8" });
            //tablaEstilos.Add("FilasSubrayadas", new string[] { "4", "6" });
            tablaEstilos = tblEstilos;

            return arrayDatos;
        }
        public static Table generarTabla(string[,] arrayDatos,Dictionary <string,string[]> confTabla)
        {
            return HerGenTabla.generarTabla(arrayDatos, confTabla);
        }
    }
}
