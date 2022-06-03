using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILicitacionesServices
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido, int idCliente);

        public Task<ResponseModel<string>> RegistrarProceso(DatoFormatoProcesoModel dato);
    }
}
