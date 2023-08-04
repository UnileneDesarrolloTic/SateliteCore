using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.Models.Encajado;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SatelliteCore.Api.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class EncajadoController : ControllerBase
    {
        private readonly IEncajadoServices _encajadoServices;

        public EncajadoController(IEncajadoServices encajadoServices)
        {
            _encajadoServices = encajadoServices;
        }

        [HttpGet("listaOrdenesFabricacion")]
        public async Task<IActionResult> ListaOrdenesFabricacion(string ordenFabricacion)
        {
            ResponseModel<List<ListaOrdenesFabricaciónDTO>> listaOrdenes = await _encajadoServices.ListaOrdenesFabricacion(ordenFabricacion);
            return Ok(listaOrdenes);
        }
    }
}
