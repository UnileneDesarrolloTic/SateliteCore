using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Request.ComprobanteOrdenCompra;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Models.Response.ComprobanteOrdenCompra;
using SatelliteCore.Api.Services;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ComprobanteOrdenCompraController : ControllerBase
    {
        private readonly IComprobanteOrdenCompraServices _comprobanteOrdenCompraServices;

        public ComprobanteOrdenCompraController(IComprobanteOrdenCompraServices comprobanteOrdenCompraServices)
        {
            _comprobanteOrdenCompraServices = comprobanteOrdenCompraServices;
        }

        [HttpGet("MostrarInformacionOrdenCompra")]
        public async Task<ActionResult> MostrarInformacionOrdenCompra(string ordenCompra, string secuencia, string item)
        {
            IEnumerable<MostrarFechaPrometida> listado = await _comprobanteOrdenCompraServices.MostrarInformacionOrdenCompra(ordenCompra, secuencia, item);

            return Ok(listado);
        }

        [HttpPost("RegistrarFechaPrometida")]
        public async Task<ActionResult> RegistrarFechaPrometida(DatosFormatoRegistrarFecha dato)
        {
            string usuario = Shared.ObtenerUsuarioSpring(HttpContext.User.Identity);
            ResponseModel<string> listado = await _comprobanteOrdenCompraServices.RegistrarFechaPrometida(dato, usuario);

            return Ok(listado);
        }

        [HttpGet("MostrarDetalleOrdenCompra")]
        public async Task<ActionResult> MostrarDetalleOrdenCompra(string ordenCompra, string item, string secuencia)
        {           
            IEnumerable<DatosFormatoDetalleOrdenCompra> listado = await _comprobanteOrdenCompraServices.MostrarDetalleOrdenCompra(ordenCompra, item, secuencia);
            return Ok(listado);
        }

    }
}
