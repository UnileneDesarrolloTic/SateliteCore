using System.Collections.Generic;

namespace SatelliteCore.Api.Models.Response.AnalisisMateriaPrima
{
    public class PlantillaProtocoloDTO
    {
        public PlantillaCabeceraProtocoloDTO Cabecera { get; set; }
        public List<PlantillaDetalleProtocoloDTO> Detalle { get; set; }
    }
}
