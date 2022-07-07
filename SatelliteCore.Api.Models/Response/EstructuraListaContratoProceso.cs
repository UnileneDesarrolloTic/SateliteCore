using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class EstructuraListaContratoProceso
    {
       public int idproceso { get; set;  }
       public string tipodeusuario { get; set; }
       public int numeroitem { get; set; }
       public string descripcionitem { get; set; }
       public  string NumeroContrato { get; set; }
    }
}
