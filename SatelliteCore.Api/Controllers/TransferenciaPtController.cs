using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.TransferenciaPT;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class TransferenciaPtController : ControllerBase
    {

        private readonly ITransferenciaPtServices _transferenciaPtServices;

        public TransferenciaPtController(ITransferenciaPtServices transferenciaPtServices)
        {
            _transferenciaPtServices = transferenciaPtServices;
        }

        [HttpGet("listaPendienteTransferenciaFisica")]
        public async Task<IActionResult> ListaPendienteTransfereciaFisica(string almacenCodigo)
        {
            ResponseModel<List<PendienteTransFisicaDTO>> listaPendientes = await _transferenciaPtServices.ListaPendienteTransfereciaFisica(almacenCodigo);
            return Ok(listaPendientes);
        }

        [HttpGet("registrarTransfenciaPT")]
        public async Task<IActionResult> RegistraTransfenciaPT(int idControl, string controlNumero, decimal cantidadTotal, decimal cantidadParcial)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);

            ResponseModel<string> result = await _transferenciaPtServices.RegistraTransfenciaPT(idControl, controlNumero, cantidadTotal,cantidadParcial, usuario);
            return Ok(result);
        }

        [HttpGet("listaPendienteRecepcionFisica")]
        public async Task<IActionResult> ListaPendienteRecepcionFisica(string almacen)
        {
            ResponseModel<List<PendienteRecepcionarPtDTO>> listaPendientes = await _transferenciaPtServices.ListaPendienteRecepcionFisica(almacen);
            return Ok(listaPendientes);
        }

        [HttpPost("registrarRecepcionPT")]
        public async Task<IActionResult> RegistraRecepcionPT(RegistrarRecepcionPtDTO recepcion)
        {
            recepcion.UsuarioRecepcion = Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);

            ResponseModel<string> result = await _transferenciaPtServices.RegistraRecepcionPT(recepcion);

            return Ok(result);
        }
    }
}
