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
        public async Task<IActionResult> ListaDetallePedido(string Pedido, int idCliente)
        {
            if (string.IsNullOrEmpty(Pedido))
                throw new ValidationModelException("El Código Pedido es Obligatorio");

            IEnumerable<ListarDetallePedido> listarDetallePedidos = await _licitacionesServices.ListaDetallePedido(Pedido, idCliente);

            return Ok(listarDetallePedidos);
        }

        [HttpPost("RegistrarProceso")]
        public async Task<IActionResult> RegistrarProceso(DatoFormatoProcesoModel dato)
        {
            ResponseModel<string> result = await _licitacionesServices.RegistrarProceso(dato);
            return Ok(result);
        }
    }
 }
