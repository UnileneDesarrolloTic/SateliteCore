using SatelliteCore.Api.Models.Entities;
using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public struct AnalisisHebraDatosGeneralesDTO
    {
        public DatosAnalisisHebraDTO Datos { get; set; }
        public TBMAnalisisHebraEntity Cabecera { get; set; }
        public List<TBDAnalisisHebraEntity> Detalle { get; set; }
    }
}
