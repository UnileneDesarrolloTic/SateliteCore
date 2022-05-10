using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class AnalisisAgujaController : ControllerBase
    {
        private readonly IAnalisisAgujaServices _analisisAgujaServices;

        public AnalisisAgujaController(IAnalisisAgujaServices analisisAgujaServices)
        {
            _analisisAgujaServices = analisisAgujaServices;
        }

        [HttpGet("ListarAnalisisAguja")]
        public async Task<IActionResult> ListarAnalisisAguja(string ordenCompra, string lote, string item, int pagina)
        {
            if (pagina == 0)
                throw new ValidationModelException("El número de página es incorrecto");

            IEnumerable<ListarAnalisisAgujaModel> listaAnalisis = await _analisisAgujaServices.ListarAnalisisAguja(ordenCompra, lote, item, pagina);

            return Ok(listaAnalisis);
        }

        [HttpGet("ListaOrdenesCompra")]
        public async Task<IActionResult> ListaOrdenesCompra(string numeroOrden)
        {
            if (string.IsNullOrEmpty(numeroOrden))
                throw new ValidationModelException("El código FOR es obligatorio");

            IEnumerable<ListarOrdenCompra> listaOrdenesCompra = await _analisisAgujaServices.ListaOrdenesCompra(numeroOrden);
            return Ok(listaOrdenesCompra);
        }

        [HttpGet("CantidadPruebasFlexionPorItem")]
        public async Task<IActionResult> CantidadPruebaFlexionPorItem(string controlNumero, int secuencia)
        {
            ResponseModel<int> cantidadPruebas = await _analisisAgujaServices.CantidadPruebaFlexionPorItem(controlNumero, secuencia);
            return Ok(cantidadPruebas);
        }

        //[HttpGet("ListarCiclos")]
        //public async Task<IActionResult> ListarCiclos(string identificador)
        //{
        //    IEnumerable<ListarAnalisisAgujaModel> listarCiclos = await _analisisAgujaServices.ListarCiclos(identificador);
        //    return Ok(listarCiclos);
        //}

        //[HttpPost("RegistrarControlAgujas")]
        //public async Task<ActionResult> RegistrarControlAgujas(ControlAgujasModel matricula)
        //{
        //    int result = await _analisisAgujaServices.RegistrarControlAgujas(matricula);

        //    ResponseModel<int> response = new ResponseModel<int>(true, "El lote se registró correctamente", result);

        //    return Ok(result);
        //}
    }
}
