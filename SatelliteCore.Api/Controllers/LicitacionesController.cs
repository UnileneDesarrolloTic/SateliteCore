using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using SatelliteCore.Api.CrossCutting.Helpers;
using SatelliteCore.Api.Models.Entities;
using SatelliteCore.Api.Models.Request;
using SatelliteCore.Api.Models.Response;
using SatelliteCore.Api.Services.Contracts;
using System.Collections.Generic;

using System.Threading.Tasks;
using SystemsIntegration.Api.Models.Exceptions;
using System.Security.Claims;


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

        [HttpGet("ObtenerTipoUsuario")]
        public async Task<IActionResult> ObtenerTipoUsuario(int NumeroProceso, int Item, string Mes)
        {
            if (NumeroProceso == 0)
                throw new ValidationModelException("El Numero de Proceso Es Obligatorio");

            IEnumerable<string> result = await _licitacionesServices.ObtenerTipoUsuario(NumeroProceso, Item, Mes);

            return Ok(result);
        }


        [HttpGet("BuscarOrdenCompraLicitaciones")]
        public async Task<IActionResult> BuscarOrdenCompraLicitaciones( int NumeroProceso, int NumeroEntrega, int Item, string TipoUsuario)
        {
            if (NumeroProceso == 0)
                throw new ValidationModelException("El Numero de Proceso Es Obligatorio");

            ResponseModel<DatosFormatoBuscarOrdenCompraLicitacionesModel> result = await _licitacionesServices.BuscarOrdenCompraLicitaciones(NumeroProceso, NumeroEntrega,Item,TipoUsuario);
            return Ok(result);
        }


        [HttpPost("RegistrarOrdenCompra")]
        public async Task<IActionResult> RegistrarOrdenCompra(DatoFormatoRegistrarOrdenCompraLicitaciones dato)
        {
            var claims = HttpContext.User.Identity as ClaimsIdentity;
            int idUsuario = int.Parse(claims.FindFirst(ClaimTypes.NameIdentifier).Value);

            ResponseModel<string> result = await _licitacionesServices.RegistrarOrdenCompra(dato, idUsuario);
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


        [HttpGet("ListarGuiaInformacion")]
        public async Task<IActionResult> ListarGuiaInformacion(string NumeroEntrega, string OrdenCompra)
        {
            if (string.IsNullOrEmpty(OrdenCompra))
                throw new ValidationModelException("Debe Ingresar la Orden de Compra");

            ResponseModel<IEnumerable<ListarGuiaInformeLPModel>> result = await _licitacionesServices.ListarGuiaInformacion(NumeroEntrega, OrdenCompra);
            return Ok(result);
        }

        [HttpGet("ListarContratoProceso")]
        public async Task<IActionResult> ListarContratoProceso(string proceso)
        {
            IEnumerable<EstructuraListaContratoProceso> result = await _licitacionesServices.ListarContratoProceso(proceso);
            return Ok(result);
        }



        [HttpPost("RegistrarContratoProceso")]
        public async Task<IActionResult> RegistrarContratoProceso(List<DatosRequestFormatoContratoProcesoModel> dato)
        {
            ResponseModel<string> result = await _licitacionesServices.RegistrarContratoProceso(dato);
            return Ok(result);
        }

        [HttpGet("DashboardLicitacionesExportar")]
        public  async Task<IActionResult> DashboardLicitacionesExportar()
        {
            ResponseModel<string> result =  await _licitacionesServices.DashboardLicitacionesExportar();
            return Ok(result);
        }
    }
}
