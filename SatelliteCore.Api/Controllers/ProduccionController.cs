using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Config;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Generic;
using SatelliteCore.Api.Models.Request;
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
    public class ProduccionController : ControllerBase
    {
        private readonly IProduccionServices _pronosticoServices;

        public ProduccionController(IProduccionServices pronosticoServices)
        {
            _pronosticoServices = pronosticoServices;
        }

        [HttpGet("ProductosArima")]
        public async Task<ActionResult> SeguimientoProductosArima(string periodo)
        {
            List<ProductoArimaModel> listaCandidatos = await _pronosticoServices.SeguimientoProductosArima(periodo);

            return Ok(listaCandidatos);
        }

        [HttpPost("ListaPedidosCreadoAuto")]
        public async Task<ActionResult> ListaPedidosCreadoAuto(PedidosCreadosDataModel filtro)
        {
            (IEnumerable<PedidosCreadosAutoLogModel> ListaPedidos, int TotalRegistros) = await _pronosticoServices.ListaPedidosCreadoAuto(filtro);

            PaginacionModel<PedidosCreadosAutoLogModel> PedidosPaginados = 
                new PaginacionModel<PedidosCreadosAutoLogModel>((List<PedidosCreadosAutoLogModel>)ListaPedidos, filtro.Pagina, filtro.RegistrosPorPagina, TotalRegistros);
                
            return Ok(PedidosPaginados);
        }

        [HttpGet("SegimientoCandidatosMP")]
        public async Task<ActionResult> SegimientoCandidatosMP(string regla)
        {
            SeguimientoCandMPAGenericModel listaCandidatos = await _pronosticoServices.ListaSeguimientoCandidatosMP(regla);
            return Ok(listaCandidatos);
        }

        [HttpGet("ControlCalidadItemMP")]
        public async Task<ActionResult> ControlCalidadItemMP(string Item)
        {
            List<DetalleControlCalidadItemMP> listaDetalleControlCalidad = await _pronosticoServices.ControlCalidadItemMP(Item);
            return Ok(listaDetalleControlCalidad);
        }

        [HttpGet("MostrarColumnaMP")]
        public async Task<ActionResult> MostrarColumnaMP()
        {
            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                             new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED, "");

                return BadRequest(responseError);
            }

            int idUsuario =  Shared.ObtenerUsuarioSesion(HttpContext.User.Identity);
            bool Permiso = await _pronosticoServices.MostrarColumnaMP(idUsuario);
            ResponseModel<dynamic> responseSuccesss = new ResponseModel<dynamic>(true, Constante.MESSAGE_SUCCESS, new {permisoColumna=Permiso});

            return Ok(responseSuccesss);
        }

        
        [HttpPost("CompraMateriaPrima")]
        public async Task<ActionResult> PronosticoCompraMP(PronosticoCompraMP dato)
        {   

            if (!ModelState.IsValid)
            {
                ResponseModel<string> responseError =
                             new ResponseModel<string>(false, Constante.MODEL_VALIDATION_FAILED,"");

                return BadRequest(responseError);
            }

            List<CompraMPArimaModel> listaProductos = await _pronosticoServices.SeguimientoCompraMPArima(dato);

            return Ok(listaProductos);
        }

    }
}