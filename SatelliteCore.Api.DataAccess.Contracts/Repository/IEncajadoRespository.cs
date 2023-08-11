using SatelliteCore.Api.Models.Encajado;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface IEncajadoRespository
    {
        public Task<List<ListaOrdenesFabricaciónDTO>> ListaOrdenesFabricacion(string ordenFabricacion, string lote);
        public Task<List<TransferenciaEncajadoDTO>> ListaTransferenciasEncaje(string ordenFabricacion);
        public Task<int> RegistarNuevaTrasnferencia(string ordenFabricacion, decimal cantidad, string usuario);
        public Task<(decimal, decimal, List<AsignacionEncajadoDTO>)> ListraAsignacionesEncajePorEtapa( int idEncaje, int etapa);
        public Task<int> RegistrarAsignacionEncaje(DatosRegistrarAsignacionDTO asignacion);
        public Task ActualizaEstadoAsignacion(int id, string estado, string usuario);
    }
}
