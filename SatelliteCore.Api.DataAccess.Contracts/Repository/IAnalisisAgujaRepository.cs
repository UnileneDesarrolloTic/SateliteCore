using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IAnalisisAgujaRepository
    {
        public Task<IEnumerable<ListarAnalisisAgujaModel>> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina);
        public Task<IEnumerable<ListarOrdenCompra>> ListaOrdenesCompra(string NumeroOrden);
        public Task<int> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia);
        //public Task<int> RegistrarControlAgujas(ControlAgujasModel matricula);
        //public Task<IEnumerable<ListarAnalisisAgujaModel>> ListarCiclos(string identificador);
    }
}
