using System;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Report.RRHH
{
    public struct AutoSobretiempoPersonaCabeceraDTO
    {
        public int IdPersona { get; set; }
        public string Nombres { get; set; }
        public string Area { get; set; }
        public string CentroCosto { get; set; }
        public string SubArea { get; set; }
    }

    public struct AutoSobretiempoPersonaDetalleDTO
    {
        public int IdPersona { get; set; }
        public DateTime FechaRegistro { get; set; }
        public string HoraInicio { get; set; }
        public string HoraFin { get; set; }
        public int Cant_horas { get; set; }
    }

    public struct AutorizacionSobretiempoPersonaDTO
    {
        public List<AutoSobretiempoPersonaCabeceraDTO> Cabecera { get; set; }
        public List<AutoSobretiempoPersonaDetalleDTO> Detalle { get; set; }
    }
}
