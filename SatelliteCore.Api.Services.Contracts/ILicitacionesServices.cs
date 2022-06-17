using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Request;
using System.Collections.Generic;
using System.Threading.Tasks;
using SatelliteCore.Api.Models.Entities;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ILicitacionesServices
    {
        public Task<IEnumerable<ListarDetallePedido>> ListaDetallePedido(string Pedido);
        public Task<ResponseModel<string>> RegistrarProceso(DatoFormatoProcesoModel dato);
        public Task<IEnumerable<DatosFormatoDistribuccionLP>> ListarDistribuccionProceso(int NumeroProceso, string Item, string Mes);
        public Task<ResponseModel<string>> RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato);
        public Task<IEnumerable<ListarProcesoEntity>> ListarProceso();
        public Task<IEnumerable<DatosFormatoProgramacionMuestraModel>> ListarProgramaMuestraLIP(int IdProceso, string NumeroEntrega);
        public Task<ResponseModel<string>> RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> dato);
    }
}
