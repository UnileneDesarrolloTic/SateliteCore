using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Request
{
    public class DatosFormatoRegistroPruebasAgujasModel
    {
        public string Especialidad { get; set; }
        public string Lote { get; set; }
        public List<GuardarPruebaFlexionAgujaModel> Detalle { get; set; }
    }
}
