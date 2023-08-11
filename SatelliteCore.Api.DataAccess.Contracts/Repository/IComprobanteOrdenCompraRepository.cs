using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra;
using SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IComprobanteOrdenCompraRepository
    {
        public Task<IEnumerable<MostrarFechaPrometida>> MostrarInformacionOrdenCompra(string ordenCompra, string secuencia, string item);
        public Task<string> RegistrarFechaPrometida(DatosFormatoRegistrarFecha dato, string usuario);
        public Task<IEnumerable<DatosFormatoDetalleOrdenCompra>> MostrarDetalleOrdenCompra(string ordenCompra, string item, string secuencia);
    }
}
