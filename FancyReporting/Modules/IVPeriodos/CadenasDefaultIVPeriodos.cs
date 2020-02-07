using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FancyReporting.Modules.IVPeriodos
{
    public static class CadenasDefaultIVPeriodos
    {
        public static Dictionary<string, string> leerCadenas
        {
            get
            {
                Dictionary<string, string> cadenas = new Dictionary<string, string>();
                //genericos
                cadenas["dato"] = "cuenta";
                cadenas["formatoFecha"] = "dd-MM-yyyy";

                cadenas["subTituloGeneral"] = "EVOLUCION DE";

                cadenas["artSingular"] = "la";
                cadenas["artPlural"] = "las";
                cadenas["fuente"] = "Fuente: Datos obtenidos de la Base de Datos Institucional.";

                //parrafos
                cadenas["asciendeSingular"] = "creció en";
                cadenas["asciendePlural"] = "crecieron en";
                cadenas["desciendeSingular"] = "decreció en";
                cadenas["desciendePlural"] = "decrecieron en";

                cadenas["tenFinAsciendeSingular"] = "tiene un desempeño mayor de";
                cadenas["tenFinAsciendePlural"] = "tienen un desempeño mayor de";
                cadenas["tenFinDesciendeSingular"] = "decreció en su desempeño en";
                cadenas["tenFinDesciendePlural"] = "decrecieron en su desempeño en";
                cadenas["ultimoConector"] = "y";
                //parrafos 4 periodos default
                cadenas["iniParIni"] = "Como se aprecia del cuadro superior,";
                cadenas["iniParMed"] = "Asimismo, del cuadro se desprende que";
                cadenas["iniParConFinal"] = "Por consiguiente, al realizar la comparación de la ejecución";

                cadenas["contraste"] = ", mientras que";
                cadenas["entreFechasFormato1"] = "entre el {0} y el {1},";
                cadenas["entreFechasFormato2"] = "desde el {0} al {1},";
                cadenas["entreFechasFormato3"] = "desde el {0} al {1}";
                cadenas["parrafoResultado"] = "del ejercicio {0} al {1} con el {2} también al mismo mes, {3}, se aprecia que en el ejercicio {2},";
                cadenas["finDescripcionPlural"] = ", respectivamente";
                cadenas["formatoParrResultado"] = "#,0.00%";

                //tablas
                cadenas["tituloPeriodos"] = "PERIODOS";
                //tabla 4 periodos
                cadenas["tituloFunciones"] = "CRECIMIENTO / DECRECIMIENTO";
                cadenas["evolucionFinal"] = "% EVOLUCIÓN";

                cadenas["formatoTblResultado"] = "#,0.00%";
                cadenas["formatoTblValPeriodo"] = "#,0.00%";

                //graficos estadisticos
                cadenas["inicioTitGraEst"] = "EVOLUCION DE ";
                cadenas["intGraEstSingular"] = "La tendencia de la evolución de {0} {1} del grupo {2} se muestra en el siguiente gráfico:";
                cadenas["intGraEstPlural"] = "Las tendencias de la evolución de {0} {1} del grupo {2} se muestran en el siguiente gráfico:";
                cadenas["formtGraficoPer"] = "MMM. yyyy";
                cadenas["formatoNumGrafico"] = "#,##0";
                cadenas["graficoEjeY"] = "MILES DE SOLES";
                //tamanos
                cadenas["anchoPrimera"] = "3000";
                cadenas["anchoPeriodos"] = "2500";
                cadenas["anchoResultados"] = "2000";

                return cadenas;
            }
        }
    }
}
