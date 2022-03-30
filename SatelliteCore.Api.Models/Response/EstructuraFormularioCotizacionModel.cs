using System;

namespace SatelliteCore.Api.Models.Response
{
    public struct EstructuraFormularioCotizacionModel
    {
        public int CodCampo { get; set; }
        public string Etiqueta { get; set; }
        public string ColumnaResp { get; set; }
        public string TipoDato { get; set; }
        public Boolean Requerido { get; set; }
        public string ValorDefecto { get; set; }
    }
}
