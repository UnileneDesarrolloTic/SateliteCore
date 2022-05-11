using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IProduccionServices
    {
        public Task<List<ProductoArimaModel>> SeguimientoProductosArima(string periodo);
        public Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro);
        public Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla);

        public Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item);

        public Task<List<CompraMPArimaModel>> SeguimientoCompraMPArima(PronosticoCompraMP dato);

    }
}
