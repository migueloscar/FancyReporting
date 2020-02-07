using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FancyReporting.Models;
using FancyReporting.Modules.UltimosPeriodos;
using FancyReporting.Modules.IVPeriodos;
using System.IO;

namespace FancyReporting
{
    public class GeneracionInforme
    {
        Informe inf;
        /// <summary>
        /// Es recomendable ejecutar esta funcion dentro de un TRY CATCH, para capturar 
        /// las excepciones del validador.
        /// </summary>
        /// <param name="modelo"></param>
        /// <param name="listaDatosJson"></param>
        /// <param name="configuracionJson"></param>
        public Informe cargarModeloInforme(TipoReporte modelo,string listaDatosJson, string configuracionJson)
        {
            
            switch (modelo)
            {
                case TipoReporte.InformeIVPeriodos://4columnas
                    {
                        inf = new GenerarInformeIVPeriodos();
                    }; break;
                case TipoReporte.UltimosPeriodos://ultimosPeriodos
                    {
                        inf = new GeneracionInformeUltimosPeriodos();
                    }; break;
                default://4columnas
                    {
                        inf = new GenerarInformeIVPeriodos();
                    }; break;
            }

            inf.inicializarInforme(listaDatosJson,configuracionJson);
            inf.agregarFuncionesPlantilla();

            return inf;
        }


        public MemoryStream generarReporte(TipoReporte modelo, string listaDatosJson, string configuracionJson)
        {
            Informe inf;
            switch (modelo)
            {
                case TipoReporte.InformeIVPeriodos://4columnas
                    {
                        inf = new GenerarInformeIVPeriodos();
                    }; break;
                case TipoReporte.UltimosPeriodos://ultimosPeriodos
                    {
                        inf = new GeneracionInformeUltimosPeriodos();
                    }; break;
                default://4columnas
                    {
                        inf = new GenerarInformeIVPeriodos();
                    }; break;
            }

            MemoryStream file = inf.generarInformeMemStream(listaDatosJson, configuracionJson);

            return file;
        }
    }
}
