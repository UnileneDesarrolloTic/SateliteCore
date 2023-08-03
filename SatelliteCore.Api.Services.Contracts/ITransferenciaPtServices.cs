using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.TransferenciaPT;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Services.Contracts
{
    public interface ITransferenciaPtServices
    {
        public Task<ResponseModel<List<PendienteTransFisicaDTO>>> ListaPendienteTransfereciaFisica(string almacenCodigo);
        public Task<ResponseModel<string>> RegistraTransfenciaPT(int idControl, string controlNumero, decimal cantidadTotal, decimal cantidadParcial, string usuario);
        public Task<ResponseModel<List<PendienteRecepcionarPtDTO>>> ListaPendienteRecepcionFisica(string almacenCodigo, string estado);
        public Task<ResponseModel<string>> RegistraRecepcionPT(RegistrarRecepcionPtDTO recepcion);
        public Task<ResponseModel<string>> ReporteTransferencia();
    }
}
