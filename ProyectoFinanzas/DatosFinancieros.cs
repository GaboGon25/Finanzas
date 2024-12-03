using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProyectoFinanzas
{
    public class DatosFinancieros
    {
        public static List<DatosBalance> BalanceGeneral { get; set; } = new List<DatosBalance>();
        public static List<DatosEstadoResultado> EstadoResultados { get; set; } = new List<DatosEstadoResultado>();
    }
}
