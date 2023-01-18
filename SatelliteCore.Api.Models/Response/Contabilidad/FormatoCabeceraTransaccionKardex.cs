using SatelliteCore.Api.Models.Generic;

namespace SatelliteCore.Api.Models.Response.Contabilidad
{
    public struct FormatoCabeceraTransaccionKardex
    {
        public decimal CCantidadTotal { get; set; }
        public decimal CMontoTotal { get; set; }
    }

    public struct InformacionTransaccionKardex
    {
        public FormatoCabeceraTransaccionKardex ContentidoCabecera { get; set; }

        public PaginacionModel<FormatoListadoInformacionTransaccionKardex> ContentidoDetalle { get; set; }
    }
}
