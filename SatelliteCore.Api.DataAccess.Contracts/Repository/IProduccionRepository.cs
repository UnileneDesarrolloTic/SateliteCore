using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IProduccionRepository
    {
        public Task<SeguimientoProductoArimaModel> SeguimientoProductosArima(string periodo);
        public Task<(IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros)> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro);
        public Task<SeguimientoCandMPAGenericModel> ListaSeguimientoCandidatosMP(string regla);

        public Task<List<DetalleControlCalidadItemMP>> ControlCalidadItemMP(string Item);
        public Task<SeguimientoComprasMPArima> SeguimientoCompraMPArima(PronosticoCompraMP dato);

    }
}
