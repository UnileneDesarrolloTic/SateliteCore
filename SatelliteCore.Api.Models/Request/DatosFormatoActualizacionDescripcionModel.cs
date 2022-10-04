using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoActualizacionDescripcionModel
    {
        public int IdDescripcion { get; set; }
        public string Marca { get; set; }
        public string Hebra { get; set; }
        public string DescripcionLocal { get; set; }
        public string DescripcionIngles { get; set; }
        public List<DatosFormatoDetalleAgujaDescripcion> DetalleAgujas { get; set; }
    }
}
