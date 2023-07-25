using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response.GestioEquipoEngaste
{
    public struct DatosFormatoInformacionEquipoEngaste
    {
        public DatosFormatoCabeceraInformacionEquipo cabecera { get; set; }
        public List<DatosFormatoDetalleInformacionEquipo> detalle { get; set; }
    }
}
