using SatelliteCore.Api.Models.Encajado;
using SatelliteCore.Api.Models.Response;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface IEncajadoServices
    {
        public Task<ResponseModel<List<ListaOrdenesFabricaciónDTO>>> ListaOrdenesFabricacion(string ordenFabricacion, string lote);
        public Task<ResponseModel<List<TransferenciaEncajadoDTO>>> ListaTransferenciasEncaje(string ordenFabricacion);
        public Task<ResponseModel<string>> RegistarNuevaTrasnferencia(string ordenFabricacion, decimal cantidad, string usuario);
        public Task<ResponseModel<object>> ListraAsignacionesEncajePorEtapa( int idEncaje, int etapa);
        public Task<ResponseModel<string>> RegistrarAsignacion(DatosRegistrarAsignacionDTO asignacion);
        public Task<ResponseModel<string>> ActualizaEstadoAsignacion(int id, string estado, string usuario);

    }
}
