using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra;
using SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IComprobanteOrdenCompraServices
    {
        public Task<IEnumerable<MostrarFechaPrometida>> MostrarInformacionOrdenCompra(string ordenCompra, string secuencia, string item);
        public Task<ResponseModel<string>> RegistrarFechaPrometida(DatosFormatoRegistrarFecha dato, string usuario);
        public Task<IEnumerable<DatosFormatoDetalleOrdenCompra>> MostrarDetalleOrdenCompra(string ordenCompra, string item, string secuencia);
    }
}
