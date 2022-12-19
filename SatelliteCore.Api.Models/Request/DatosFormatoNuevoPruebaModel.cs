using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoNuevoPruebaModel
    {
        public int IdMedologia { get; set;  }
        public int IdAgrupadoHebra { get; set; }
        public string IdCalibre { get; set; }
        public string IdUnidadMedida { get; set; }
        public string valor { get; set; }
        public string DescripcionLocal { get; set; }
        public string DescripcionIngle { get; set; }
        public string EspecificacionLocal { get; set; }
        public string EspecificacionIngles { get; set; }
    }
}
