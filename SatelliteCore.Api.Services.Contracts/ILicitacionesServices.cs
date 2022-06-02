using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILicitacionesServices
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido, int idCliente);
    }
}
