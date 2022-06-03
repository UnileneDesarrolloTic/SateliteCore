using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ILicitacionesRepository
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido,int idCliente);

        public Task<int> RegistrarProceso(DatoFormatoProcesoModel matricula);
    }
}
