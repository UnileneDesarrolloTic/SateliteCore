using SatelliteCore.Api.DataAccess.Contracts.Repository;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.TransferenciaPT;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Services
{
    public class TransferenciaPtServices : ITransferenciaPtServices
    {
        private readonly ITransferenciaPtRepository _transferenciaPtRepository;
        public TransferenciaPtServices(ITransferenciaPtRepository transferenciaPtRepository)
        {
            _transferenciaPtRepository = transferenciaPtRepository;
        }

        public async Task<ResponseModel<List<PendienteTransFisicaDTO>>> ListaPendienteTransfereciaFisica(string almacenCodigo)
        {
            if(string.IsNullOrWhiteSpace(almacenCodigo))
                throw new ValidationModelException();

            List<PendienteTransFisicaDTO> lista = await _transferenciaPtRepository.ListaPendienteTransfereciaFisica(almacenCodigo);

            return new ResponseModel<List<PendienteTransFisicaDTO>>(lista);
        }

        public async Task<ResponseModel<string>> RegistraTransfenciaPT(int idControl, string controlNumero, decimal cantidadTotal, decimal cantidadParcial, string usuario)
        {
            if (idControl == 0 || string.IsNullOrWhiteSpace(controlNumero) || string.IsNullOrEmpty(usuario) || cantidadParcial == (decimal) 0.0 || cantidadParcial == (decimal)0.0)
                throw new ValidationModelException();

            await _transferenciaPtRepository.RegistraTransfenciaPT(idControl, controlNumero, cantidadTotal, cantidadParcial, usuario);

            return new ResponseModel<string>("Se ha registrado la transferencia.");
        }

        public async Task<ResponseModel<List<PendienteRecepcionarPtDTO>>> ListaPendienteRecepcionFisica(string almacenCodigo, string estado)
        {
            if(string.IsNullOrWhiteSpace(almacenCodigo) || string.IsNullOrWhiteSpace(estado))
                throw new ValidationModelException();

            List<PendienteRecepcionarPtDTO> lista = await _transferenciaPtRepository.ListaPendienteRecepcionFisica(almacenCodigo, estado);

            return new ResponseModel<List<PendienteRecepcionarPtDTO>>(lista);
        }

        public async Task<ResponseModel<string>> RegistraRecepcionPT(RegistrarRecepcionPtDTO recepcion)
        {
            if(!recepcion.ValidarDatos())
                throw new ValidationModelException();

            string mensaje = await _transferenciaPtRepository.RegistraRecepcionPT(recepcion);

            return new ResponseModel<string>(mensaje);

        }

    }
}
