using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class LicitacionesController : ControllerBase
    {
        private readonly ILicitacionesServices _licitacionesServices;
        public LicitacionesController(ILicitacionesServices licitacionesServices)
        {
            _licitacionesServices = licitacionesServices;
        }

        [HttpGet("ListaDetallePedido")]
        public async Task<IActionResult> ListaDetallePedido(string Pedido)
        {
            if (string.IsNullOrEmpty(Pedido))
                throw new ValidationModelException("El Código Pedido es Obligatorio");

            IEnumerable<ListarDetallePedido> listarDetallePedidos = await _licitacionesServices.ListaDetallePedido(Pedido);

            return Ok(listarDetallePedidos);
        }

        [HttpPost("RegistrarProceso")]
        public async Task<IActionResult> RegistrarProceso(DatoFormatoProcesoModel dato)
        {
            ResponseModel<string> result = await _licitacionesServices.RegistrarProceso(dato);
            return Ok(result);
        }

        [HttpGet("ListarDistribuccionProceso")]
        public async Task<IActionResult> ListarDistribuccionProceso(int NumeroProceso, string Item , string Mes)
        {
            if (NumeroProceso==0)
                throw new ValidationModelException("El Numero de Proceso Es Obligatorio");

            IEnumerable<DatosFormatoDistribuccionLP> result = await _licitacionesServices.ListarDistribuccionProceso(NumeroProceso, Item, Mes);
            
            return Ok(result);
        }
        

        [HttpPost("RegistrarDistribuccionProceso")]
        public async Task<IActionResult> RegistrarDistribuccionProceso(List<DatoFormatoDistribuccionLPModel> dato)
        {
            ResponseModel<string> result = await _licitacionesServices.RegistrarDistribuccionProceso(dato);

            return Ok(result);
        }

        [HttpGet("ListarProceso")]
        public async Task<IActionResult> ListarProceso()
        {
            IEnumerable<ListarProcesoEntity> result = await _licitacionesServices.ListarProceso();

            return Ok(result);
        }

        [HttpGet("ListarProgramaMuestraLIP")]
        public async Task<IActionResult> ListarProgramaMuestraLIP(int IdProceso, string NumeroEntrega)
        {
            IEnumerable<DatosFormatoProgramacionMuestraModel> result = await _licitacionesServices.ListarProgramaMuestraLIP(IdProceso, NumeroEntrega);
            return Ok(result);
        }


        [HttpPost("RegistrarProgramacionMuestreo")]
        public async Task<IActionResult> RegistrarProgramacionMuestreo(List<DatosFormatoMuestraEnsayoLIP> dato)
        {
            ResponseModel<string> result = await _licitacionesServices.RegistrarProgramacionMuestreo(dato);
            return Ok(result);
        }


    }
}
