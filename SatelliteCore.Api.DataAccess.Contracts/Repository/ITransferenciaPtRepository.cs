using SatelliteCore.Api.Models.Response.TransferenciaPT;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.DataAccess.Contracts.Repository
{
    public interface ITransferenciaPtRepository
    {
        public Task<List<PendienteTransFisicaDTO>> ListaPendienteTransfereciaFisica(string almacenCodigo);
        public Task RegistraTransfenciaPT(int idControl, string controlNumero, decimal cantidadTotal, decimal cantidadParcial, string usuario);
        public Task<List<PendienteRecepcionarPtDTO>> ListaPendienteRecepcionFisica(string almacenCodigo, string estado);
        public Task<string> RegistraRecepcionPT(RegistrarRecepcionPtDTO recepcion);
        public Task<List<DatosRptTransferenciaPT>> DatosReporteTransferencia();


    }
}
