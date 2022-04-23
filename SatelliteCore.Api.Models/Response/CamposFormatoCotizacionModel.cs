using System;
using System.Collections.Generic;
using System.Text;

namespace SatelliteCore.Api.Models.Response
{
    public class CamposFormatoCotizacionModel
    {

        public int IDFormato { get; set; }
        public string TipoDetalle { get; set; }
        public int CodCampo { get; set; }
        public string Etiqueta { get; set; }
        public string ColumnaResp { get; set; }
        public string TipoDatoTs { get; set; }
        public int Requerido { get; set; }
        public string ValorDefecto { get; set; }
        public string ColumnaScript { get; set; }

    }
}
